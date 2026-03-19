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

    // --- Excel生成 (xlsx-populate) ---
    try {
      const XlsxPopulate = require("xlsx-populate");
      const workbook = await XlsxPopulate.fromFileAsync(templatePath);
      
      // 不要なシートを削除 (申込書のみ残す)
      const targetSheetName = "申込書";
      workbook.sheets().forEach((s: any) => {
        if (s.name() !== targetSheetName) s.delete();
      });
      let sheet = workbook.sheet(targetSheetName);
      if (!sheet) sheet = workbook.sheet(0);

      // --- ヘッダー文字配置とスタイリング ---
      const style13 = { fontSize: 13 };
      const style14 = { fontSize: 14 };
      
      sheet.cell("B1").value("【訪問理美容サービス　申込書】").style({ fontSize: 23, bold: true });
      sheet.cell("B3").value("下記の項目をご記入のうえ、お申し込みください。").style(style13);
      sheet.cell("B4").value("※ご希望の施術内容に〇印をつけてください。").style(style13);
      sheet.cell("B5").value("※「施術開始時間の希望」がございましたら、第一～第三希望までご記入ください。").style(style13);
      sheet.cell("B6").value("※お申込みはご訪問日の5日前までとなります。").style(style13);
      sheet.cell("B7").value("※顔そり、シャンプーのみのご注文を承っておりません。").style(style13);
      sheet.cell("B8").value("※カラー初回の方はパッチテストを実施し、異常なければ翌月から実施となります。（パッチテストに費用はかかりません）").style(style13);
      sheet.cell("B9").value("※施術終了後にサインをいただき、施術実施月の月末までにデータにて事務局まで送付ください。（こちらをもとに御請求書を発行いたします）").style(style13);
      sheet.cell("B10").value("事務局アドレス").style(style13);
      sheet.cell("D10").value("houmon@hannan-plage.jp").style(style13);

      // 施設名・施術日 (I列, N列, O列)
      const yearInt = parseInt(year || "0");
      const reiwa = yearInt > 2018 ? yearInt - 2018 : "";
      sheet.cell("I3").value(`施設名：${facilityName || ""}`).style(style14);
      sheet.cell("N3").value("施術日：").style(style14);
      sheet.cell("O3").value(`令和 ${reiwa} 年 ${month || "  "} 月 ${day || "  "} 日`).style(style14);

      // --- 確認サイン欄 (動的) ---
      const customers = body.customers || [];
      const hasGuider = customers.some((c: CustomerData) => c.isGuided && String(c.isGuided).trim() !== "" && String(c.isGuided).trim() !== "否");
      
      // サイン欄の描画
      if (hasGuider) {
        sheet.range("S1:U1").merged(true).value("確認サイン欄").style({ horizontalAlignment: "center", fill: "e0e0e0", border: true });
        sheet.cell("S2").value("施設側").style({ border: true, horizontalAlignment: "center" });
        sheet.cell("T2").value("案内担当").style({ border: true, horizontalAlignment: "center" });
        sheet.cell("U2").value("訪問側").style({ border: true, horizontalAlignment: "center" });
        sheet.range("S3:U3").value("サイン").style({ border: true, horizontalAlignment: "center", fontSize: 9 });
        sheet.range("S4:S9").merged(true).style({ border: true });
        sheet.range("T4:T9").merged(true).style({ border: true });
        sheet.range("U4:U9").merged(true).style({ border: true });
      } else {
        sheet.range("S1:T1").merged(true).value("確認サイン欄").style({ horizontalAlignment: "center", fill: "e0e0e0", border: true });
        sheet.cell("S2").value("施設側").style({ border: true, horizontalAlignment: "center" });
        sheet.cell("T2").value("訪問側").style({ border: true, horizontalAlignment: "center" });
        sheet.range("S3:T3").value("サイン").style({ border: true, horizontalAlignment: "center", fontSize: 9 });
        sheet.range("S4:S9").merged(true).style({ border: true });
        sheet.range("T4:T9").merged(true).style({ border: true });
      }

      // --- レイアウトの動적検出 (データ行) ---
      let rowHeader = 13, rowTotal = 14, rowExample = 15, rowDataStart = 16;
      for (let r = 10; r <= 30; r++) {
        const val = String(sheet.row(r).cell(2).value() || "");
        if (val.includes("No.")) rowHeader = r;
        if (val.includes("合計人数")) rowTotal = r;
        if (val.includes("記入例")) { rowExample = r; rowDataStart = r + 1; }
      }

      // メニュー列の特定
      const menuCols: Record<string, number> = {};
      for (const r of [rowHeader - 1, rowHeader]) {
        for (let c = 7; c <= 15; c++) {
          const v = String(sheet.row(r).cell(c).value() || "");
          if (v && !["No.", "部屋番号", "氏名", "性別", "合計料金", "施術開始時間の希望", "合計", "料金"].some(s => v.includes(s))) {
            const cleanV = v.replace(/\n/g, "").split(" ")[0].split("(")[0].trim();
            if (cleanV) menuCols[cleanV] = c;
          }
        }
      }

      // --- 顧客データ書き込み ---
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

        row.cell(16).value(customer.isGuided || "否");
        if (customer.hasService) row.cell(17).value("サイン");
        row.cell(18).value(customer.isAdditionalMenuAllowed || "可・否");
        row.cell(19).value(customer.isCustomOrder || "本人・お任せ");
        row.cell(20).value(customer.remarks || "");
      });

      // --- 合計人数行 ---
      if (rowTotal) {
        sheet.row(rowTotal).cell(4).value(customers.length);
      }

      const buffer = await workbook.outputAsync();

      const formatDate = `${year || ""}${month || ""}${day || ""}`;
      const fileName = `${facilityName || "申込書"}_${formatDate || "清書"}.xlsx`
        .replace(/[/\\?%*:|"<>]/g, "_");

      return new NextResponse(buffer, {
        status: 200,
        headers: {
          "Content-Disposition": `attachment; filename="${encodeURIComponent(fileName)}"`,
          "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        },
      });
    } catch (innerError: any) {
      console.error("Inner Excel Error:", innerError);
      return NextResponse.json({ error: "Excel生成中にエラーが発生しました", details: innerError.message }, { status: 500 });
    }
  } catch (error: unknown) {
    console.error("API Route Error:", error);
    const message = error instanceof Error ? error.message : "Unknown error";
    return NextResponse.json({ error: "Excel生成処理中にエラーが発生しました", details: message }, { status: 500 });
  }
}
