import { NextResponse } from "next/server";
import ExcelJS from "exceljs";
import path from "path";
import fs from "fs";

// --- 型定義 ---
interface CustomerData {
  no: number;
  room: string;
  name: string;
  gender?: string;
  menus: Record<string, boolean>;
  total_price: number;
  is_cancelled: boolean;
  remarks: string;
  time_slots?: string[];
}

interface ExportRequest {
  facility_name: string;
  date: string;
  headers: string[];          // ["氏名", "カット", "カラー(白髪染め)", ..., "施術実施"]
  prices: Record<string, number>;  // { "カット": 2000, "カラー": 4000, ... }
  customers: CustomerData[];
}

// --- 定数 ---
const HEADER_ROW_1 = 12;  // カテゴリヘッダー行
const HEADER_ROW_2 = 13;  // メニュー名行
const PRICE_ROW = 14;     // 単価行
const TOTAL_ROW = 15;     // 合計人数行
const DATA_START_ROW = 17; // データ開始行

// 固定列
const COL_NO = 2;        // B列: No.
const COL_ROOM = 3;      // C列: 部屋番号
const COL_NAME = 4;      // D列: 氏名 (D-Eは結合)
const MENU_START_COL = 6; // F列からメニュー開始

// 施設名セル・日付セル
const FACILITY_CELL = "H3";
const DATE_CELL = "O3";

export async function POST(req: Request) {
  try {
    const body: ExportRequest = await req.json();
    const { facility_name, date, headers, prices, customers } = body;

    // メニュー列の特定（"氏名"と"施術実施"を除外）
    const menuHeaders = headers.filter(
      (h) => h !== "氏名" && h !== "施術実施"
    );
    const menuCount = menuHeaders.length;

    // 合計金額列・施術開始時間列・備考列の位置を計算
    const colTotal = MENU_START_COL + menuCount;      // メニューの次の列
    const colTime1 = colTotal + 1;                     // 施術開始時間 第一希望
    const colTime2 = colTotal + 2;                     // 第二希望
    const colTime3 = colTotal + 3;                     // 第三希望
    const colStatus = colTotal + 4;                    // ご案内 有無
    const colDone = colTotal + 5;                      // 施術実施 有無
    const colRemarks = colTotal + 6;                   // 備考

    // --- ベーステンプレートを読み込み ---
    const templatePath = path.join(
      process.cwd(),
      "public",
      "templates",
      "digitizer_base.xlsx"
    );

    const workbook = new ExcelJS.Workbook();

    if (fs.existsSync(templatePath)) {
      await workbook.xlsx.readFile(templatePath);
    } else {
      // テンプレートがない場合はゼロから作成
      workbook.addWorksheet("申込書");
    }

    const worksheet = workbook.worksheets[0];
    worksheet.name = "申込書";

    // --- 施設名・日付の書き込み ---
    worksheet.getCell(FACILITY_CELL).value = facility_name || "";
    worksheet.getCell(DATE_CELL).value = date || "";

    // --- ヘッダー行の構築 ---
    // 行12: カテゴリヘッダー
    const row12 = worksheet.getRow(HEADER_ROW_1);
    row12.getCell(COL_NO).value = "No.";
    row12.getCell(COL_ROOM).value = "部屋番号";
    row12.getCell(COL_NAME).value = "氏名";
    for (let i = 0; i < menuCount; i++) {
      row12.getCell(MENU_START_COL + i).value = "メニュー/料金";
    }
    row12.getCell(colTotal).value = "合計料金";
    row12.getCell(colTime1).value = "施術開始時間の希望";
    row12.getCell(colStatus).value = "ご案内\n有無";
    row12.getCell(colDone).value = "施術実施\n有無";
    row12.getCell(colRemarks).value = "備考";

    // 行13: メニュー名
    const row13 = worksheet.getRow(HEADER_ROW_2);
    row13.getCell(COL_NO).value = "No.";
    row13.getCell(COL_ROOM).value = "部屋番号";
    row13.getCell(COL_NAME).value = "氏名";
    for (let i = 0; i < menuCount; i++) {
      row13.getCell(MENU_START_COL + i).value = menuHeaders[i];
    }
    row13.getCell(colTotal).value = "合計金額";
    row13.getCell(colTime1).value = "第一希望";
    row13.getCell(colTime2).value = "第二希望";
    row13.getCell(colTime3).value = "第三希望";
    row13.getCell(colStatus).value = "ご案内\n有無";
    row13.getCell(colDone).value = "施術実施\n有無";
    row13.getCell(colRemarks).value = "備考";

    // 行14: 単価
    const row14 = worksheet.getRow(PRICE_ROW);
    row14.getCell(COL_NO).value = "No.";
    row14.getCell(COL_ROOM).value = "部屋番号";
    row14.getCell(COL_NAME).value = "氏名";
    for (let i = 0; i < menuCount; i++) {
      const menuName = menuHeaders[i];
      row14.getCell(MENU_START_COL + i).value = prices[menuName] ?? 0;
    }
    row14.getCell(colTotal).value = "合計金額";

    // ヘッダー行のスタイル
    for (const row of [row12, row13, row14]) {
      row.eachCell({ includeEmpty: false }, (cell) => {
        cell.font = { bold: true, size: 9 };
        cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
        cell.border = {
          top: { style: "thin" },
          bottom: { style: "thin" },
          left: { style: "thin" },
          right: { style: "thin" },
        };
      });
    }

    // --- 合計行 (行15) ---
    const row15 = worksheet.getRow(TOTAL_ROW);
    row15.getCell(COL_NO).value = "合計人数";
    const lastDataRow = DATA_START_ROW + Math.max(customers.length - 1, 0);
    // 氏名列のCOUNTA
    const nameColLetter = getColumnLetter(COL_NAME);
    row15.getCell(COL_NAME).value = {
      formula: `COUNTA(${nameColLetter}${DATA_START_ROW}:${nameColLetter}${lastDataRow + 30})`,
    };
    // メニュー列のCOUNTIF
    for (let i = 0; i < menuCount; i++) {
      const col = MENU_START_COL + i;
      const colLetter = getColumnLetter(col);
      row15.getCell(col).value = {
        formula: `COUNTIF(${colLetter}${DATA_START_ROW}:${colLetter}${lastDataRow + 30},"〇")`,
      };
    }
    // 合計金額のSUM
    const totalColLetter = getColumnLetter(colTotal);
    row15.getCell(colTotal).value = {
      formula: `SUM(${totalColLetter}${DATA_START_ROW}:${totalColLetter}${lastDataRow + 30})`,
    };
    row15.getCell(COL_NO).font = { bold: true, size: 9 };

    // --- データ行の書き込み ---
    customers.forEach((customer, index) => {
      const rowNum = DATA_START_ROW + index;
      const row = worksheet.getRow(rowNum);

      // No.
      row.getCell(COL_NO).value = customer.no || index + 1;
      // 部屋番号
      row.getCell(COL_ROOM).value = customer.room || "";
      // 氏名
      row.getCell(COL_NAME).value = customer.name || "";

      // メニュー列
      for (let i = 0; i < menuCount; i++) {
        const menuName = menuHeaders[i];
        if (customer.is_cancelled) {
          row.getCell(MENU_START_COL + i).value = "";
        } else {
          row.getCell(MENU_START_COL + i).value = customer.menus[menuName] ? "〇" : "";
        }
      }

      // 合計金額（数式で計算 or 直接値）
      if (customer.is_cancelled) {
        row.getCell(colTotal).value = 0;
      } else {
        // 数式: メニュー列の〇がある列の単価を合計
        const menuFormulaParts: string[] = [];
        for (let i = 0; i < menuCount; i++) {
          const menuColLetter = getColumnLetter(MENU_START_COL + i);
          const priceColLetter = getColumnLetter(MENU_START_COL + i);
          menuFormulaParts.push(
            `IF(${menuColLetter}${rowNum}="〇",${priceColLetter}${PRICE_ROW},0)`
          );
        }
        if (menuFormulaParts.length > 0) {
          row.getCell(colTotal).value = {
            formula: menuFormulaParts.join("+"),
          };
        } else {
          row.getCell(colTotal).value = customer.total_price || 0;
        }
      }

      // 施術開始時間
      if (customer.time_slots && customer.time_slots.length > 0) {
        if (customer.time_slots[0]) row.getCell(colTime1).value = customer.time_slots[0];
        if (customer.time_slots[1]) row.getCell(colTime2).value = customer.time_slots[1];
        if (customer.time_slots[2]) row.getCell(colTime3).value = customer.time_slots[2];
      }

      // 備考
      const remarkParts: string[] = [];
      if (customer.is_cancelled) remarkParts.push("キャンセル");
      if (customer.remarks) remarkParts.push(customer.remarks);
      row.getCell(colRemarks).value = remarkParts.join(" / ") || "";

      // セルスタイル
      row.eachCell({ includeEmpty: false }, (cell) => {
        cell.font = { size: 9 };
        cell.alignment = { vertical: "middle", horizontal: "center" };
        cell.border = {
          top: { style: "thin" },
          bottom: { style: "thin" },
          left: { style: "thin" },
          right: { style: "thin" },
        };
      });

      // 氏名列は左寄せ
      row.getCell(COL_NAME).alignment = { vertical: "middle", horizontal: "left" };

      // キャンセル行のスタイル
      if (customer.is_cancelled) {
        row.eachCell({ includeEmpty: false }, (cell) => {
          cell.font = {
            size: 9,
            strike: true,
            color: { argb: "FF999999" },
          };
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFF0F0F0" },
          };
        });
      }
    });

    // --- 列幅の設定 ---
    worksheet.getColumn(COL_NO).width = 5;
    worksheet.getColumn(COL_ROOM).width = 10;
    worksheet.getColumn(COL_NAME).width = 18;
    worksheet.getColumn(COL_NAME + 1).width = 2; // E列（氏名結合分）
    for (let i = 0; i < menuCount; i++) {
      worksheet.getColumn(MENU_START_COL + i).width = 10;
    }
    worksheet.getColumn(colTotal).width = 10;
    worksheet.getColumn(colTime1).width = 10;
    worksheet.getColumn(colTime2).width = 10;
    worksheet.getColumn(colTime3).width = 10;
    worksheet.getColumn(colStatus).width = 8;
    worksheet.getColumn(colDone).width = 8;
    worksheet.getColumn(colRemarks).width = 25;

    // --- バッファ生成・レスポンス返却 ---
    const buffer = await workbook.xlsx.writeBuffer();

    const fileName = `${facility_name || "申込書"}_${date || "清書"}.xlsx`
      .replace(/[/\\?%*:|"<>]/g, "_");

    return new NextResponse(buffer, {
      status: 200,
      headers: {
        "Content-Disposition": `attachment; filename="${encodeURIComponent(fileName)}"`,
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      },
    });
  } catch (error: unknown) {
    console.error("Excel Export Error:", error);
    const message = error instanceof Error ? error.message : "Unknown error";
    return NextResponse.json(
      { error: "Excel生成に失敗しました", details: message },
      { status: 500 }
    );
  }
}

// --- ヘルパー: 列番号をExcel列文字に変換 (1=A, 2=B, ..., 27=AA) ---
function getColumnLetter(colNum: number): string {
  let result = "";
  let n = colNum;
  while (n > 0) {
    n--;
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26);
  }
  return result;
}
