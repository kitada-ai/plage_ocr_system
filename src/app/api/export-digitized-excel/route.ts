import { NextResponse } from "next/server";
import ExcelJS from "exceljs";

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
  headers: string[];
  prices: Record<string, number>;
  customers: CustomerData[];
}

// --- 定数 ---
const TITLE_ROW = 1;
const FACILITY_ROW = 3;
const HEADER_ROW = 5;   // メニュー名ヘッダー
const PRICE_ROW = 6;    // 単価行
const DATA_START_ROW = 7; // データ開始行

// 固定列
const COL_NO = 1;        // A列: No.
const COL_ROOM = 2;      // B列: 部屋番号
const COL_NAME = 3;      // C列: 氏名
const MENU_START_COL = 4; // D列からメニュー開始

export async function POST(req: Request) {
  try {
    const body: ExportRequest = await req.json();
    const { facility_name, date, headers, prices, customers } = body;

    // メニュー列の特定（"氏名"と"施術実施"を除外）
    const menuHeaders = headers.filter(
      (h) => h !== "氏名" && h !== "施術実施"
    );
    const menuCount = menuHeaders.length;

    // 動的列位置の計算
    const colTotal = MENU_START_COL + menuCount;      // 合計金額
    const colRemarks = colTotal + 1;                   // 備考

    // --- ワークブックをゼロから作成 ---
    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet("申込書");

    // --- 列幅の設定 ---
    ws.getColumn(COL_NO).width = 5;
    ws.getColumn(COL_ROOM).width = 10;
    ws.getColumn(COL_NAME).width = 16;
    for (let i = 0; i < menuCount; i++) {
      ws.getColumn(MENU_START_COL + i).width = 12;
    }
    ws.getColumn(colTotal).width = 10;
    ws.getColumn(colRemarks).width = 25;

    // --- スタイル定義 ---
    const thinBorder: Partial<ExcelJS.Borders> = {
      top: { style: "thin" },
      bottom: { style: "thin" },
      left: { style: "thin" },
      right: { style: "thin" },
    };

    const headerFill: ExcelJS.Fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFD9E2F3" },
    };

    const priceFill: ExcelJS.Fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFF2CC" },
    };

    // --- 行1: タイトル ---
    const titleCell = ws.getCell(TITLE_ROW, 1);
    titleCell.value = "【訪問理美容サービス 申込書】";
    titleCell.font = { bold: true, size: 14 };
    ws.mergeCells(TITLE_ROW, 1, TITLE_ROW, Math.min(colRemarks, 8));

    // --- 行3: 施設名・日付 ---
    const facCell = ws.getCell(FACILITY_ROW, 1);
    facCell.value = `施設名：${facility_name || ""}`;
    facCell.font = { bold: true, size: 11 };
    ws.mergeCells(FACILITY_ROW, 1, FACILITY_ROW, 3);

    const dateCell = ws.getCell(FACILITY_ROW, MENU_START_COL);
    dateCell.value = `施術日：${date || ""}`;
    dateCell.font = { bold: true, size: 11 };
    ws.mergeCells(FACILITY_ROW, MENU_START_COL, FACILITY_ROW, Math.min(MENU_START_COL + 3, colRemarks));

    // --- 行5: ヘッダー（メニュー名）---
    const headerRow = ws.getRow(HEADER_ROW);
    headerRow.height = 30;
    const headerDefs = [
      { col: COL_NO, value: "No." },
      { col: COL_ROOM, value: "部屋番号" },
      { col: COL_NAME, value: "氏名" },
      ...menuHeaders.map((h, i) => ({ col: MENU_START_COL + i, value: h })),
      { col: colTotal, value: "合計金額" },
      { col: colRemarks, value: "備考" },
    ];
    for (const def of headerDefs) {
      const cell = headerRow.getCell(def.col);
      cell.value = def.value;
      cell.font = { bold: true, size: 9 };
      cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
      cell.border = thinBorder;
      cell.fill = headerFill;
    }

    // --- 行6: 単価行 ---
    const priceRow = ws.getRow(PRICE_ROW);
    priceRow.height = 20;
    const pCell1 = priceRow.getCell(COL_NO);
    pCell1.value = "";
    pCell1.border = thinBorder;
    const pCell2 = priceRow.getCell(COL_ROOM);
    pCell2.value = "";
    pCell2.border = thinBorder;
    const pCell3 = priceRow.getCell(COL_NAME);
    pCell3.value = "（単価）";
    pCell3.font = { size: 8, italic: true };
    pCell3.alignment = { horizontal: "center", vertical: "middle" };
    pCell3.border = thinBorder;
    pCell3.fill = priceFill;
    for (let i = 0; i < menuCount; i++) {
      const cell = priceRow.getCell(MENU_START_COL + i);
      const menuName = menuHeaders[i];
      const price = prices[menuName] ?? 0;
      cell.value = price > 0 ? `¥${price.toLocaleString()}` : "¥0";
      cell.font = { size: 8 };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = thinBorder;
      cell.fill = priceFill;
    }
    const pCellTotal = priceRow.getCell(colTotal);
    pCellTotal.value = "";
    pCellTotal.border = thinBorder;
    pCellTotal.fill = priceFill;
    const pCellRemarks = priceRow.getCell(colRemarks);
    pCellRemarks.value = "";
    pCellRemarks.border = thinBorder;
    pCellRemarks.fill = priceFill;

    // --- データ行の書き込み ---
    customers.forEach((customer, index) => {
      const rowNum = DATA_START_ROW + index;
      const row = ws.getRow(rowNum);
      row.height = 22;

      // No.
      const cellNo = row.getCell(COL_NO);
      cellNo.value = customer.no || index + 1;
      cellNo.alignment = { horizontal: "center", vertical: "middle" };
      cellNo.border = thinBorder;

      // 部屋番号
      const cellRoom = row.getCell(COL_ROOM);
      cellRoom.value = customer.room || "";
      cellRoom.alignment = { horizontal: "center", vertical: "middle" };
      cellRoom.border = thinBorder;

      // 氏名
      const cellName = row.getCell(COL_NAME);
      cellName.value = customer.name || "";
      cellName.alignment = { horizontal: "left", vertical: "middle" };
      cellName.border = thinBorder;

      // メニュー列
      for (let i = 0; i < menuCount; i++) {
        const menuName = menuHeaders[i];
        const cell = row.getCell(MENU_START_COL + i);
        if (customer.is_cancelled) {
          cell.value = "";
        } else {
          cell.value = customer.menus[menuName] ? "〇" : "";
        }
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.border = thinBorder;
        cell.font = { size: 10 };
      }

      // 合計金額（数式で計算）
      const cellTotal = row.getCell(colTotal);
      if (customer.is_cancelled) {
        cellTotal.value = 0;
      } else {
        const formulaParts: string[] = [];
        for (let i = 0; i < menuCount; i++) {
          const menuColLetter = getColumnLetter(MENU_START_COL + i);
          const priceColLetter = getColumnLetter(MENU_START_COL + i);
          formulaParts.push(
            `IF(${menuColLetter}${rowNum}="〇",${priceColLetter}${PRICE_ROW},0)`
          );
        }
        if (formulaParts.length > 0) {
          cellTotal.value = { formula: formulaParts.join("+") };
        } else {
          cellTotal.value = customer.total_price || 0;
        }
      }
      cellTotal.alignment = { horizontal: "right", vertical: "middle" };
      cellTotal.border = thinBorder;
      cellTotal.numFmt = "¥#,##0";

      // 備考
      const cellRemarks = row.getCell(colRemarks);
      const remarkParts: string[] = [];
      if (customer.is_cancelled) remarkParts.push("キャンセル");
      if (customer.remarks) remarkParts.push(customer.remarks);
      cellRemarks.value = remarkParts.join(" / ") || "";
      cellRemarks.alignment = { horizontal: "left", vertical: "middle", wrapText: true };
      cellRemarks.border = thinBorder;

      // キャンセル行のスタイル
      if (customer.is_cancelled) {
        for (let c = COL_NO; c <= colRemarks; c++) {
          const cell = row.getCell(c);
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
        }
      }
    });

    // --- 合計行 ---
    const lastDataRow = DATA_START_ROW + Math.max(customers.length - 1, 0);
    const totalRowNum = lastDataRow + 1;
    const totalRow = ws.getRow(totalRowNum);
    totalRow.height = 22;

    const totalLabel = totalRow.getCell(COL_NO);
    totalLabel.value = "合計";
    totalLabel.font = { bold: true, size: 10 };
    totalLabel.alignment = { horizontal: "center", vertical: "middle" };
    totalLabel.border = thinBorder;
    totalLabel.fill = headerFill;
    ws.mergeCells(totalRowNum, COL_NO, totalRowNum, COL_NAME);

    // メニューごとのCOUNTIF
    for (let i = 0; i < menuCount; i++) {
      const col = MENU_START_COL + i;
      const colLetter = getColumnLetter(col);
      const cell = totalRow.getCell(col);
      cell.value = {
        formula: `COUNTIF(${colLetter}${DATA_START_ROW}:${colLetter}${lastDataRow},"〇")`,
      };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = thinBorder;
      cell.fill = headerFill;
      cell.font = { bold: true, size: 10 };
    }

    // 合計金額のSUM
    const totalColLetter = getColumnLetter(colTotal);
    const sumCell = totalRow.getCell(colTotal);
    sumCell.value = {
      formula: `SUM(${totalColLetter}${DATA_START_ROW}:${totalColLetter}${lastDataRow})`,
    };
    sumCell.alignment = { horizontal: "right", vertical: "middle" };
    sumCell.border = thinBorder;
    sumCell.fill = headerFill;
    sumCell.font = { bold: true, size: 10 };
    sumCell.numFmt = "¥#,##0";

    // 備考セル（空）
    const remarksTotal = totalRow.getCell(colRemarks);
    remarksTotal.value = "";
    remarksTotal.border = thinBorder;
    remarksTotal.fill = headerFill;

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

// --- ヘルパー: 列番号をExcel列文字に変換 ---
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
