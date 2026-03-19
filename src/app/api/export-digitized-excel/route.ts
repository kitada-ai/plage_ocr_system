import { NextResponse } from "next/server";
import path from "path";
import fs from "fs";
import os from "os";

// --- 型定義 ---
interface MenuItem {
  name: string;
  price: number;
}

interface CustomerData {
  no: number;
  room: string;
  name: string;
  gender: string;
  selectedMenus: string[];
  preferredTimes: string[];
  hasService: boolean;
  isGuided: string;
  isAdditionalMenuAllowed: string;
  isCustomOrder: string;
  remarks: string;
}

interface ExportRequest {
  facilityName: string;
  year: string;
  month: string;
  day: string;
  dayOfWeek: string;
  menuItems: MenuItem[];
  customers: CustomerData[];
}

export async function POST(req: Request) {
  try {
    const body = await req.json();
    const facilityName = body.facilityName;
    const dateStr = body.date || ""; 
    const [year, month, day] = dateStr.split("-");

    const writtenTemplatesDir = path.join(process.cwd(), "docs", "samples", "written");
    let templatePath = path.join(process.cwd(), "docs", "samples", "sheets", "【★ﾄﾞｸﾀｰｻﾝｺﾞ守口様】訪問施術サービス （申込書・請求書）.xlsx");

    try {
      if (fs.existsSync(writtenTemplatesDir)) {
        const files = fs.readdirSync(writtenTemplatesDir);
        const match = files.find((f: string) => f.endsWith(".xlsx") && facilityName && f.includes(facilityName));
        if (match) {
          templatePath = path.join(writtenTemplatesDir, match);
          console.log(`Using facility-specific template: ${templatePath}`);
        }
      }
    } catch (e) {
      console.error("Template selection error:", e);
    }

    if (!fs.existsSync(templatePath)) {
      const fallbackPath = path.join(process.cwd(), "docs", "samples", "sheets", "【★ﾄﾞｸﾀｰｻﾝｺﾞ守口様】訪問施術サービス （申込書・請求書）.xlsx");
      if (fs.existsSync(fallbackPath)) {
        templatePath = fallbackPath;
      } else {
        return NextResponse.json({ error: "テンプレートが見つかりません" }, { status: 500 });
      }
    }

    // --- Excel生成 ---
    try {
      const XlsxPopulate = require("xlsx-populate");
      const workbook = await XlsxPopulate.fromFileAsync(templatePath);
      
      // 「申込書」シート以外を削除
      workbook.sheets().forEach((s: any) => {
        if (s.name() !== "申込書") s.delete();
      });
      const sheet = workbook.sheet("申込書");

      // --- レイアウトの動的検出 ---
      let rowFacility = 3, colFacility = 9, colDate = 16, rowHeader = 13, rowTotal = 15, rowExample = 16, rowDataStart = 17;
      for (let r = 1; r <= 30; r++) {
        for (let c = 1; c <= 25; c++) {
          const val = String(sheet.row(r).cell(c).value() || "");
          if (val.includes("施設名")) { rowFacility = r; colFacility = c; }
          if (val.includes("施術日")) { colDate = c; }
          if (val.includes("No.")) { rowHeader = r; }
          if (val.includes("合計人数")) { rowTotal = r; }
          if (val.includes("記入例")) { rowExample = r; rowDataStart = r + 1; }
        }
      }

      // --- ヘッダー・注意書きの精密設定 (文字サイズ・内容) ---
      sheet.cell("B1").value("【訪問理美容サービス　申込書】").style({ fontSize: 23, bold: true });
      sheet.cell("B3").value("下記の項目をご記入のうえ、お申し込みください。").style("fontSize", 13);
      sheet.cell("B4").value("※ご希望の施術内容に〇印をつけてください。").style("fontSize", 13);
      sheet.cell("B5").value("※「施術開始時間の希望」がございましたら、第一～第三希望までご記入ください。").style("fontSize", 13);
      sheet.cell("B6").value("※お申込みはご訪問日の5日前までとなります。").style("fontSize", 13);
      sheet.cell("B7").value("※顔そり、シャンプーのみのご注文を承っておりません。").style("fontSize", 13);
      sheet.cell("B8").value("※カラー初回の方はパッチテストを実施し、異常なければ翌月から実施となります。（パッチテストに費用はかかりません）").style("fontSize", 13);
      sheet.cell("B9").value("※施術終了後にサインをいただき、施術実施月の月末までにデータにて事務局まで送付ください。（こちらをもとに御請求書を発行いたします）").style("fontSize", 13);
      sheet.cell("B10").value("事務局アドレス").style("fontSize", 13);
      sheet.cell("D10").value("houmon@hannan-plage.jp").style("fontSize", 13);

      // 右上 (施設名・施術日)
      const yearInt = parseInt(year || "0");
      const reiwa = yearInt > 2018 ? yearInt - 2018 : "  ";
      sheet.cell("I3").value(`施設名：${facilityName || ""}`).style("fontSize", 14);
      sheet.cell("N3").value("施術日：").style("fontSize", 14);
      sheet.cell("O3").value(`令和 ${reiwa} 年 ${month || "  "} 月 ${day || "  "} 日`).style("fontSize", 14);

      // 下線などのスタイリング
      sheet.range("I3:M3").style({ border: { bottom: true } });
      sheet.range("O3:R3").style({ border: { bottom: true } });

      // --- 確認サイン欄 ---
      const customers = body.customers || [];
      const hasGuide = customers.some((c: any) => c.isGuided && String(c.isGuided).trim() !== "");
      
      const signRange = hasGuide ? "S1:U9" : "S1:T9";
      const signBox = sheet.range(signRange);
      signBox.style({ border: true, fill: "f2f2f2" });
      
      sheet.range(hasGuide ? "S1:U2" : "S1:T2").merged(true).value("確認サイン欄").style({ horizontalAlignment: "center", verticalAlignment: "center", bold: true });
      
      const signHeaders = hasGuide ? ["施設側", "案内担当", "訪問側"] : ["施設側", "訪問側"];
      signHeaders.forEach((text, i) => {
        const col = String.fromCharCode(83 + i); 
        sheet.cell(`${col}3`).value(text).style({ horizontalAlignment: "center", fontSize: 11 });
        sheet.cell(`${col}4`).value("サイン").style({ horizontalAlignment: "center", fontSize: 9, fontColor: "808080" });
        sheet.range(`${col}5:${col}9`).style({ border: { left: true, right: true, top: true, bottom: true }, fill: "ffffff" });
      });

      // --- データ書き込み ---
      const menuCols: Record<string, number> = {};
      const searchRow = rowHeader || 11;
      for (let c = 7; c <= 15; c++) {
        const v = String(sheet.row(searchRow).cell(c).value() || sheet.row(searchRow - 1).cell(c).value() || "");
        if (v && !["No.", "部屋番号", "氏名", "性別", "合計料金", "施術開始時間の希望", "合計", "料金"].some(s => v.includes(s))) {
          const cleanV = v.replace(/\n/g, "").split(" ")[0].split("(")[0].trim();
          if (cleanV) menuCols[cleanV] = c;
        }
      }

      customers.forEach((customer: CustomerData, idx: number) => {
        const rowNum = rowDataStart + idx;
        const row = sheet.row(rowNum);
        
        row.cell(2).value(customer.no || (idx + 1));
        row.cell(3).value(customer.room || "");
        row.cell(4).value(customer.name || "");
        row.cell(6).value(customer.gender || "");

        const selectedMenus = customer.selectedMenus || [];
        selectedMenus.forEach((mName: string) => {
          for (const [colName, colIdx] of Object.entries(menuCols)) {
            if (mName.includes(colName) || colName.includes(mName)) {
              row.cell(colIdx).value("〇");
              break;
            }
          }
        });

        const times = customer.preferredTimes || [];
        times.forEach((t: string, i: number) => { if (i < 3) row.cell(13 + i).value(t); });

        row.cell(16).value(customer.isGuided || "");
        if (customer.hasService) row.cell(17).value("サイン");
        row.cell(18).value(customer.isAdditionalMenuAllowed || "可・否");
        row.cell(19).value(customer.isCustomOrder || "本人・お任せ");
        row.cell(20).value(customer.remarks || "");
      });

      // --- 合計人数行 (記入例の上) ---
      if (rowTotal) {
        let totalValCol = 4;
        for (let c = 1; c <= 10; c++) {
          const val = String(sheet.row(rowTotal).cell(c).value() || "");
          if (val.includes("合計人数")) { totalValCol = c + 1; break; }
        }
        sheet.row(rowTotal).cell(totalValCol).value(customers.length).style({ bold: true });
      }

      const buffer = await workbook.outputAsync();

      const formatDate = `${year || ""}${month || ""}${day || ""}`;
      const fileName = `${facilityName || "申込書"}_${formatDate || "清書"}.xlsx`
        .replace(/[/\\?%*:|"<>]/g, "_");

      return new NextResponse(buffer as any, {
        status: 200,
        headers: {
          "Content-Disposition": `attachment; filename="${encodeURIComponent(fileName)}"`,
          "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        },
      });
    } catch (error: any) {
      console.error("Excel Generation Error:", error);
      return NextResponse.json(
        { error: "Excel生成に失敗しました", details: error.message },
        { status: 500 }
      );
    }
  } catch (error: unknown) {
    console.error("API Route Error:", error);
    const message = error instanceof Error ? error.message : "Unknown error";
    return NextResponse.json(
      { error: "Excel生成処理中にエラーが発生しました", details: message },
      { status: 500 }
    );
  }
}
