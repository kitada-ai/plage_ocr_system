import { NextResponse } from "next/server";
import path from "path";
import fs from "fs";
import os from "os";
import ExcelJS from "exceljs";

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
  isGuided?: string;
  isAdditionalMenuAllowed?: string;
  isCustomOrder?: string;
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
    const body: ExportRequest = await req.json();
    const { facilityName, year, month, day } = body;

    // テンプレートファイルの選択
    const writtenTemplatesDir = path.join(process.cwd(), "docs", "samples", "written");
    let templatePath = path.join(
      writtenTemplatesDir,
      "0203改訂【ドクターサンゴ守口】訪問施術サービス申込書v3.xlsx"
    );

    // 施設名に一致するファイル名を探す
    try {
      if (fs.existsSync(writtenTemplatesDir)) {
        const files = fs.readdirSync(writtenTemplatesDir);
        const match = files.find((f: string) => f.endsWith(".xlsx") && facilityName && f.includes(facilityName));
        if (match) {
          templatePath = path.join(writtenTemplatesDir, match);
          console.log(`Using facility-specific template: ${match}`);
        }
      }
    } catch (e) {
      console.error("Template selection error:", e);
    }

    if (!fs.existsSync(templatePath)) {
      // フォールバック（これもない場合はエラー）
      const fallbackPath = path.join(process.cwd(), "docs", "samples", "sheets", "【★ﾄﾞｸﾀｰｻﾝｺﾞ守口様】訪問施術サービス （申込書・請求書）.xlsx");
      if (fs.existsSync(fallbackPath)) {
        templatePath = fallbackPath;
      } else {
        return NextResponse.json(
          { error: "テンプレートファイルが見つかりません" },
          { status: 500 }
        );
      }
    }

    // --- Excel生成 ---
    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(templatePath);
      const worksheet = workbook.worksheets[0];

      // --- レイアウトの動的検出 ---
      let rowFacility = 3, colFacility = 9, colDate = 16, rowHeader = 13, rowTotal = 14, rowExample = 15, rowDataStart = 16;
      for (let r = 1; r <= 21; r++) {
        const row = worksheet.getRow(r);
        for (let c = 1; c <= 25; c++) {
          const val = String(row.getCell(c).value || "");
          if (val.includes("施設名")) { rowFacility = r; colFacility = c; }
          if (val.includes("施術日")) { colDate = c; }
          if (val.includes("No.")) { rowHeader = r; }
          if (val.includes("合計人数")) { rowTotal = r; }
          if (val.includes("記入例")) { rowExample = r; rowDataStart = r + 1; }
        }
      }

      // --- ヘッダー書き込み ---
      worksheet.getRow(rowFacility).getCell(colFacility).value = `施設名：${facilityName || ""}`;
      const yearInt = parseInt(year || "0");
      const reiwa = yearInt > 2018 ? yearInt - 2018 : "";
      worksheet.getRow(rowFacility).getCell(colDate).value = `施術日：令和 ${reiwa} 年 ${month || "  "} 月 ${day || "  "} 日`;

      // --- メニュー列の特定 ---
      const menuCols: Record<string, number> = {};
      for (const r of [rowHeader - 1, rowHeader]) {
        const row = worksheet.getRow(r);
        for (let c = 7; c <= 14; c++) {
          const v = String(row.getCell(c).value || "");
          if (v && !["No.", "部屋番号", "氏名", "性別", "合計料金", "施術開始時間の希望", "合計", "料金"].some(s => v.includes(s))) {
            const cleanV = v.replace(/\n/g, "").split(" ")[0].split("(")[0].trim();
            if (cleanV) menuCols[cleanV] = c;
          }
        }
      }

      // --- 顧客データ書き込み ---
      const customers = body.customers || [];
      customers.forEach((customer, idx) => {
        const rowNum = rowDataStart + idx;
        const row = worksheet.getRow(rowNum);
        const exampleRow = worksheet.getRow(rowExample);

        // 基本情報
        row.getCell(2).value = customer.no || (idx + 1);
        row.getCell(3).value = customer.room || "";
        row.getCell(4).value = customer.name || "";
        row.getCell(6).value = customer.gender || "";

        // スタイルのコピー (記入例から)
        for (let c = 2; c <= 20; c++) {
          const srcCell = exampleRow.getCell(c);
          const dstCell = row.getCell(c);
          dstCell.style = { ...srcCell.style };
        }

        // メニュー選択 (〇)
        const selectedMenus = customer.selectedMenus || [];
        selectedMenus.forEach(mName => {
          for (const [colName, colIdx] of Object.entries(menuCols)) {
            if (mName.includes(colName) || colName.includes(mName)) {
              row.getCell(colIdx).value = "〇";
              break;
            }
          }
        });

        // 希望時間
        const times = customer.preferredTimes || [];
        times.forEach((t, i) => { if (i < 3) row.getCell(13 + i).value = t; });

        // その他
        row.getCell(16).value = customer.isGuided || "";
        if (customer.hasService) row.getCell(17).value = "サイン";
        row.getCell(18).value = customer.isAdditionalMenuAllowed || "可・否";
        row.getCell(19).value = customer.isCustomOrder || "本人・お任せ";
        row.getCell(20).value = customer.remarks || "";
      });

      // --- 合計人数行 ---
      if (rowTotal) {
        const totalRow = worksheet.getRow(rowTotal);
        totalRow.getCell(4).value = customers.length;
        for (let c = 6; c <= 14; c++) {
          if (totalRow.getCell(c).value !== null) {
            const ltr = String.fromCharCode(64 + c);
            totalRow.getCell(c).value = { formula: `COUNTIF(${ltr}${rowDataStart}:${ltr}${rowDataStart + customers.length + 5},"〇")` };
          }
        }
      }

      const buffer = await workbook.xlsx.writeBuffer();

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
    } catch (excelError: any) {
      console.error("Excel Generation Error:", excelError);
      return NextResponse.json(
        { error: "Excel生成に失敗しました", details: excelError.message },
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

