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
        if (match) templatePath = path.join(writtenTemplatesDir, match);
      }
    } catch (e) {
      console.error("Template selection error:", e);
    }

    if (!fs.existsSync(templatePath)) {
      const fallbackPath = path.join(process.cwd(), "docs", "samples", "sheets", "【★ﾄﾞｸﾀｰｻﾝｺﾞ守口様】訪問施術サービス （申込書・請求書）.xlsx");
      if (fs.existsSync(fallbackPath)) templatePath = fallbackPath;
      else return NextResponse.json({ error: "テンプレートが見つかりません" }, { status: 500 });
    }

    // --- Excel生成 ---
    try {
      const XlsxPopulate = require("xlsx-populate");
      const workbook = await XlsxPopulate.fromFileAsync(templatePath);
      
      workbook.sheets().forEach((s: any) => {
        if (s.name() !== "申込書") s.delete();
      });
      const sheet = workbook.sheet("申込書");

      // --- 1. 基本設定・列幅・フォント ---
      sheet.range("A1:Z500").style("fontName", "Yu Gothic");
      
      const colWidths: Record<string, number> = {
        B: 5, C: 10, D: 18, E: 6, F: 12, G: 12, H: 12, I: 12, J: 12, K: 12,
        L: 12, M: 12, N: 12, O: 12, P: 12, Q: 12, R: 14, S: 14, T: 35
      };
      Object.entries(colWidths).forEach(([col, width]) => {
        sheet.column(col).width(width);
      });

      // --- 2. スタイル・罫線・背景色 (Row 12 - 46) ---
      // 固定枠のクリアと罫線設定
      sheet.range("B11:T500").value(null).style({ border: false, fill: false });
      
      const gridRange = sheet.range("B12:T46");
      gridRange.style({
        border: { top: true, bottom: true, left: true, right: true, style: 'thin', color: '000000' },
        verticalAlignment: "center"
      });

      // 背景色指定
      sheet.range("B12:T14").style("fill", "E0E0E0"); // ヘッダー
      sheet.range("B15:T15").style("fill", "FFF2CC"); // 合計
      sheet.range("B16:T16").style("fill", "DDEBF7"); // 記入例

      // アライメント設定
      sheet.range("B12:T16").style("horizontalAlignment", "center");
      sheet.range("B17:T46").style("horizontalAlignment", "center"); // デフォルト中央
      sheet.range("D17:D46").style("horizontalAlignment", "left");   // 氏名: 左
      sheet.range("T17:T46").style("horizontalAlignment", "left");   // 備考: 左
      sheet.range("L17:L46").style("horizontalAlignment", "right");  // 料金: 右

      // --- 3. 共通テキスト (1-11行目) ---
      sheet.cell("B1").value("【訪問理美容サービス　申込書】").style({ fontSize: 23, bold: true });
      sheet.cell("B3").value("下記の項目をご記入のうえ、お申し込みください。").style("fontSize", 13);
      sheet.cell("B4").value("※ご希望の施術内容に〇印をつけてください。").style("fontSize", 13);
      sheet.cell("B5").value("※「施術開始時間の希望」がございましたら、第一～第三希望までご記入ください。").style("fontSize", 13);
      
      // OCR抽出の結果（注意書き）
      const customers = body.customers || [];
      const noticeLines = body.noticeLines || []; // 施設特有の注意書き
      noticeLines.forEach((line: string, i: number) => {
        if (i < 3) sheet.cell(`B${6 + i}`).value(line).style("fontSize", 13);
      });
      
      sheet.cell("B9").value("※施術終了後にサインをいただき、月末までにデータにて送付ください。").style("fontSize", 13);
      sheet.cell("B10").value("事務局アドレス").style("fontSize", 13);
      sheet.cell("D10").value("houmon@hannan-plage.jp").style("fontSize", 13);

      // --- 4. 動的・抽出テキスト (ヘッダー) ---
      const yearInt = parseInt(year || "0");
      const reiwa = yearInt > 2018 ? yearInt - 2018 : "  ";
      sheet.cell("I3").value(`施設名： ${facilityName || ""}`).style({ fontSize: 14, underline: true });
      sheet.cell("N3").value(`施術日： 令和 ${reiwa} 年 ${month || "  "} 月 ${day || "  "} 日`).style({ fontSize: 14, underline: true });

      sheet.cell("P2").value("▼確認サイン▼").style({ horizontalAlignment: "center" });
      const signHeaders = body.signHeaders || ["施設側", "訪問側"]; // デフォルト
      signHeaders.forEach((text: string, i: number) => {
        if (i < 3) {
          const col = String.fromCharCode(80 + i); // P, Q, R
          sheet.cell(`${col}3`).value(text).style({ border: true, horizontalAlignment: "center" });
          sheet.range(`${col}4:${col}9`).style({ border: true });
        }
      });

      // 表見出し (12-14行目)
      sheet.cell("B12").value("No.");
      sheet.cell("C12").value("部屋番号");
      sheet.cell("D12").value("氏名");
      sheet.cell("L12").value("合計料金");
      sheet.cell("M12").value("施術開始時間の希望");
      sheet.cell("P12").value("ご案内有無");
      sheet.cell("Q12").value("施術実施有無");
      sheet.cell("T12").value("備考");
      sheet.cell("M14").value("第一希望");
      sheet.cell("N14").value("第二希望");
      sheet.cell("O14").value("第三希望");

      const menuItems = body.menuItems || [];
      const hasGender = customers.some((c: any) => c.gender && c.gender.trim() !== "");
      if (hasGender) sheet.cell("E12").value("性別");

      const hasAddMenu = customers.some((c: any) => c.isAdditionalMenuAllowed && c.isAdditionalMenuAllowed.trim() !== "");
      if (hasAddMenu) sheet.cell("R12").value("追加メニュー可否");

      const hasCustomOrder = customers.some((c: any) => c.isCustomOrder && c.isCustomOrder.trim() !== "");
      if (hasCustomOrder) sheet.cell("S12").value("オーダーメイド");

      sheet.cell("M13").value(body.timeInstruction || "※30分単位等");

      // メニューと料金の動的配置 (F13-K14)
      menuItems.forEach((item: any, i: number) => {
        if (i < 6) {
          const col = String.fromCharCode(70 + i); // F, G, H, I, J, K
          sheet.cell(`${col}13`).value(item.name || "");
          sheet.cell(`${col}14`).value(item.price ? `¥${item.price.toLocaleString()}` : "");
        }
      });

      // --- 5. 集計・記入例・実データ ---
      sheet.cell("B15").value("合計人数").style("bold", true);
      sheet.cell("B16").value("記入例");

      // 合計人数の算出 (D列に集計)
      sheet.cell("D15").value(customers.length).style("bold", true);
      // メニューごとの合計 (F-K列)
      for (let i = 0; i < Math.min(6, menuItems.length); i++) {
        const col = String.fromCharCode(70 + i);
        const count = customers.filter((c: any) => c.selectedMenus && c.selectedMenus.includes(menuItems[i].name)).length;
        sheet.cell(`${col}15`).value(count || "");
      }

      // 実データ流し込み
      const rowDataStart = 17;
      customers.forEach((customer: CustomerData, idx: number) => {
        if (idx >= 30) return; // 46行目まで
        const rowNum = rowDataStart + idx;
        const row = sheet.row(rowNum);
        
        row.cell(2).value(customer.no || (idx + 1));
        row.cell(3).value(customer.room || "");
        row.cell(4).value(customer.name || "");
        row.cell(6).value(customer.gender || "");

        // メニュー〇付け
        menuItems.forEach((m: any, mIdx: number) => {
          if (mIdx < 6) {
            if (customer.selectedMenus && customer.selectedMenus.includes(m.name)) {
              row.cell(6 + mIdx).value("〇");
            }
          }
        });

        row.cell(12).value(customer.hasService ? (customer.selectedMenus ? "" : "") : ""); // 合計料金(OCR)
        // 合計料金が必要な場合はここに
        
        const times = customer.preferredTimes || [];
        times.forEach((t: string, i: number) => { if (i < 3) row.cell(13 + i).value(t); });

        row.cell(16).value(customer.isGuided || "");
        if (customer.hasService) row.cell(17).value("✓");
        row.cell(18).value(customer.isAdditionalMenuAllowed || "");
        row.cell(19).value(customer.isCustomOrder || "");
        row.cell(20).value(customer.remarks || "");
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
