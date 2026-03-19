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

    // --- Excel生成 (xlsx-populate: 値のみ書き込み) ---
    try {
      const XlsxPopulate = require("xlsx-populate");
      const workbook = await XlsxPopulate.fromFileAsync(templatePath);
      
      // 対象シートの特定
      let sheet = workbook.sheet("申込書");
      if (!sheet) sheet = workbook.sheet(0);

      // --- レイアウトの動的検出 ---
      let rowFacility = 3, colFacility = 9, colDate = 16, rowHeader = 13, rowTotal = 14, rowExample = 15, rowDataStart = 16;
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

      // --- ヘッダー書き込み ---
      sheet.row(rowFacility).cell(colFacility).value(`施設名：${facilityName || ""}`);
      const yearInt = parseInt(year || "0");
      const reiwa = yearInt > 2018 ? yearInt - 2018 : "";
      sheet.row(rowFacility).cell(colDate).value(`施術日：令和 ${reiwa} 年 ${month || "  "} 月 ${day || "  "} 日`);

      // --- メニュー列の特定 ---
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
      const customers = body.customers || [];
      customers.forEach((customer: CustomerData, idx: number) => {
        const rowNum = rowDataStart + idx;
        const row = sheet.row(rowNum);
        
        // 既存の行を活かしつつ値を書き込む
        row.cell(2).value(customer.no || (idx + 1));
        row.cell(3).value(customer.room || "");
        row.cell(4).value(customer.name || "");
        row.cell(6).value(customer.gender || "");

        // メニュー選択 (〇)
        const selectedMenus = customer.selectedMenus || [];
        selectedMenus.forEach((mName: string) => {
          for (const [colName, colIdx] of Object.entries(menuCols)) {
            if (mName.includes(colName) || colName.includes(mName)) {
              row.cell(colIdx).value("〇");
              break;
            }
          }
        });

        // 希望時間
        const times = customer.preferredTimes || [];
        times.forEach((t: string, i: number) => { if (i < 3) row.cell(13 + i).value(t); });

        // その他
        row.cell(16).value(customer.isGuided || "");
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

