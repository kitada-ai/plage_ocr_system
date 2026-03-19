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

      // --- ユーザー指定の絶対番地設定 ---
      // Row 1: タイトル (23pt, 太文字)
      sheet.cell("B1").value("【訪問理美容サービス　申込書】").style({ fontSize: 23, bold: true });

      // Row 3: 各種情報 (14pt)
      const yearInt = parseInt(year || "0");
      const reiwa = yearInt > 2018 ? yearInt - 2018 : "  ";
      sheet.cell("I3").value(`施設名：${facilityName || ""}`).style("fontSize", 14);
      sheet.cell("N3").value("施術日：").style("fontSize", 14);
      sheet.cell("O3").value(`令和 ${reiwa} 年 ${month || "  "} 月 ${day || "  "} 日`).style("fontSize", 14);
      sheet.range("I3:M3").style({ border: { bottom: true } });
      sheet.range("O3:R3").style({ border: { bottom: true } });

      // Row 3 (続き): 注意書き 13pt
      sheet.cell("B3").value("下記の項目をご記入のうえ、お申し込みください。").style("fontSize", 13);
      sheet.cell("B4").value("※ご希望の施術内容に〇印をつけてください。").style("fontSize", 13);
      sheet.cell("B5").value("※「施術開始時間の希望」は第一～第三希望まですべてご記入ください。").style("fontSize", 13);
      sheet.cell("B6").value("※お申込みはご訪問日の5日前までとなります。").style("fontSize", 13);
      sheet.cell("B7").value("※顔そり、シャンプーのみのご注文を承っておりません。").style("fontSize", 13);
      sheet.cell("B8").value("※カラー初回の方はパッチテストを実施し、異常なければ翌月から実施となります。").style("fontSize", 13);
      sheet.cell("B9").value("※施術終了後にサインをいただき事務局まで送付ください。").style("fontSize", 13);
      sheet.cell("B10").value("事務局アドレス").style("fontSize", 13);
      sheet.cell("D10").value("houmon@hannan-plage.jp").style("fontSize", 13);

      // --- 確認サイン欄 (S1:U9) ---
      const customers = body.customers || [];
      const hasGuide = customers.some((c: any) => c.isGuided && String(c.isGuided).trim() !== "");
      const signRange = hasGuide ? "S1:U9" : "S1:T9";
      sheet.range(signRange).style({ border: true, fill: "f2f2f2" });
      sheet.range(hasGuide ? "S1:U2" : "S1:T2").merged(true).value("確認サイン欄").style({ horizontalAlignment: "center", verticalAlignment: "center", bold: true });
      const signHeaders = hasGuide ? ["施設側", "案内担当", "訪問側"] : ["施設側", "訪問側"];
      signHeaders.forEach((text, i) => {
        const col = String.fromCharCode(83 + i); 
        sheet.cell(`${col}3`).value(text).style({ horizontalAlignment: "center", fontSize: 11 });
        sheet.cell(`${col}4`).value("サイン").style({ horizontalAlignment: "center", fontSize: 9, fontColor: "808080" });
        sheet.range(`${col}5:${col}9`).style({ border: { left: true, right: true, top: true, bottom: true }, fill: "ffffff" });
      });

      // --- ユーザー指定の行マッピング ---
      // 行11: 空白 (既存の記入例などをクリア)
      sheet.row(11).clear();

      // 行12-14: 項目のヘッダー (テンプレートのRow 8-10からコピー)
      // 他のテンプレートに合わせて動的に取得する代わりに、12-14を固定ヘッダーとする
      // 実際には 11行目以降を一度すべてクリアして構築
      for (let r = 11; r <= 100; r++) sheet.row(r).clear();

      // ヘッダー labels (Row 12, 13, 14)
      sheet.row(12).cell(2).value("No.");
      sheet.row(12).cell(3).value("部屋番号");
      sheet.row(12).cell(4).value("氏名");
      sheet.row(12).cell(6).value("性別");
      sheet.row(12).cell(7).value("メニュー／料金").style({ horizontalAlignment: "center" });
      sheet.row(13).cell(7).value("カット");
      sheet.row(13).cell(8).value("カラー");
      sheet.row(13).cell(9).value("パーマ");
      sheet.row(13).cell(10).value("マニキュア");
      sheet.row(13).cell(11).value("顔そり");
      sheet.row(13).cell(12).value("シャンプー");
      sheet.row(12).cell(13).value("合計料金");
      sheet.row(12).cell(14).value("施術開始時間の希望");
      sheet.row(14).cell(14).value("第一希望");
      sheet.row(14).cell(15).value("第二希望");
      sheet.row(14).cell(16).value("第三希望");
      sheet.row(12).cell(17).value("施術実施\n有無");
      sheet.row(12).cell(18).value("備考");

      // 行15: 合計人数
      sheet.row(15).cell(3).value("合計人数").style({ bold: true, horizontalAlignment: "right", fill: "e2f0d9" });
      sheet.row(15).cell(4).value(customers.length).style({ bold: true, fill: "e2f0d9" });

      // 行16: 記入例
      sheet.row(16).cell(2).value("記入例").style({ italic: true, fill: "f2f2f2" });
      sheet.row(16).cell(4).value("山田　太郎").style({ italic: true, fill: "f2f2f2" });
      sheet.row(16).cell(7).value("〇").style({ fill: "f2f2f2" });

      // 行17以降: 実データ
      const rowDataStart = 17;
      customers.forEach((customer: CustomerData, idx: number) => {
        const rowNum = rowDataStart + idx;
        const row = sheet.row(rowNum);
        row.cell(2).value(customer.no || (idx + 1));
        row.cell(3).value(customer.room || "");
        row.cell(4).value(customer.name || "");
        row.cell(6).value(customer.gender || "");

        // メニュー
        if (customer.selectedMenus) {
          customer.selectedMenus.forEach(m => {
            if (m.includes("カット")) row.cell(7).value("〇");
            if (m.includes("カラー")) row.cell(8).value("〇");
            if (m.includes("パーマ")) row.cell(9).value("〇");
            if (m.includes("マニキュア")) row.cell(10).value("〇");
            if (m.includes("顔そり")) row.cell(11).value("〇");
            if (m.includes("シャンプー")) row.cell(12).value("〇");
          });
        }
        
        const times = customer.preferredTimes || [];
        times.forEach((t: string, i: number) => { if (i < 3) row.cell(14 + i).value(t); });

        if (customer.hasService) row.cell(17).value("✓");
        row.cell(18).value(customer.remarks || "");
      });

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
      return NextResponse.json({ error: "Excel生成に失敗しました", details: error.message }, { status: 500 });
    }
  } catch (error: unknown) {
    console.error("API Route Error:", error);
    const message = error instanceof Error ? error.message : "Unknown error";
    return NextResponse.json({ error: "Excel生成処理中にエラーが発生しました", details: message }, { status: 500 });
  }
}
