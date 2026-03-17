import { NextResponse } from "next/server";
import ExcelJS from "exceljs";
import path from "path";
import fs from "fs";

/** 請求書テンプレートのファイル名 */
const INVOICE_TEMPLATE_FILENAME = "請求書テンプレート.xlsx";

// テンプレート構造
// B6:N7 タイトル「御請求書」, B9:F10 施設名, K2:N2 日付
// B14:C15 ご請求金額ラベル, D14:G15 合計金額表示
// K11:O14 会社名・住所・登録番号（テンプレートのまま）
// E18:E19〜J18:J19 メニュー名（2行）, E20〜J20 単価
// データ: 21-22, 23-24, 25-26, 27-28, 29-30 が各1件（2行結合）
// 総合計: K31:L32 に数式 SUM(K21:L30)

type DateData = {
  counts: Record<string, number>;
  unitPrices?: Record<string, number>;
  totalDiscount?: number;
  remarks?: string;
  facility?: string;
  reiwaYear?: string;
  month?: string;
  day?: string;
  weekday?: string;
};

export async function POST(req: Request) {
  try {
    const payload: { dateDataList: DateData[]; allowedMenus?: string[] } = await req.json();
    const { dateDataList, allowedMenus } = payload;

    if (!dateDataList || dateDataList.length === 0) {
      return NextResponse.json({ error: "データがありません" }, { status: 400 });
    }

    const templatePath = path.join(process.cwd(), "public", "templates", INVOICE_TEMPLATE_FILENAME);
    if (!fs.existsSync(templatePath)) {
      return NextResponse.json({ error: "テンプレートファイルが存在しません" }, { status: 500 });
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);
    const sheet = workbook.worksheets[0];

    const firstDate = dateDataList[0];
    const { reiwaYear, month: firstMonth } = firstDate;
    // 施設名: 最初に施設名が入っている日付のものを使用
    const facility = dateDataList.map((d) => d.facility).find((f) => f && f.trim())?.trim() || "(施設名未取得)";

    // 基本情報: 日付は K2
    const now = new Date();
    const currentReiwa = now.getFullYear() - 2018;
    sheet.getCell("K2").value = `令和${currentReiwa}年${now.getMonth() + 1}月${now.getDate()}日`;

    // タイトル: B6:N7（6・7行）
    if (reiwaYear && firstMonth) {
      sheet.getCell("B6").value = `御請求書(${reiwaYear}${firstMonth}月度)`;
    }

    // 施設名: B9:F10（9B〜10Fの結合セル）
    sheet.getCell("B9").value = facility;
    (sheet.getCell("B9") as ExcelJS.Cell).alignment = { horizontal: "center", vertical: "middle" };

    // メニュー一覧
    const menuOrderSet = new Set<string>();
    const allowedSet = allowedMenus && allowedMenus.length > 0 ? new Set(allowedMenus) : null;
    for (const d of dateDataList) {
      for (const name of [...Object.keys(d.counts || {}), ...Object.keys(d.unitPrices || {})]) {
        if (!allowedSet || allowedSet.has(name)) menuOrderSet.add(name);
      }
    }
    const menuOrder = Array.from(menuOrderSet).slice(0, 6);

    const mergedUnitPrices: Record<string, number> = {};
    for (const d of dateDataList) {
      for (const [name, price] of Object.entries(d.unitPrices || {})) {
        const p = Number(price);
        if (mergedUnitPrices[name] === undefined) mergedUnitPrices[name] = isNaN(p) ? 0 : p;
      }
    }

    // ヘッダー: 単価
    const headerCols = ["E", "F", "G", "H", "I", "J"];
    for (let c = 0; c < menuOrder.length; c++) {
      const menuName = menuOrder[c];
      const price = Number(mergedUnitPrices[menuName]) || 0;
      const cell18 = sheet.getCell(`${headerCols[c]}18`);
      cell18.value = menuName;
      (cell18 as ExcelJS.Cell).alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      const cell20 = sheet.getCell(`${headerCols[c]}20`);
      cell20.value = price;
      (cell20 as ExcelJS.Cell).numFmt = price > 0 ? '"￥"#,##0' : "";
      (cell20 as ExcelJS.Cell).alignment = { horizontal: "center", vertical: "middle" };
    }

    // データ行
    const dataRows = [21, 23, 25, 27, 29];
    let grandTotal = 0;

    for (let i = 0; i < Math.min(dateDataList.length, dataRows.length); i++) {
      const row = dataRows[i];
      const { counts, unitPrices = {}, month, day, weekday, totalDiscount = 0, remarks = "" } = dateDataList[i];
      const prices = { ...mergedUnitPrices, ...unitPrices };

      // 月・日
      const monthStr = month != null && String(month).trim() !== "" ? String(month) : "";
      const dayStr = day != null && String(day).trim() !== "" ? String(day) : "";
      const monthCell = sheet.getCell(row, 2);
      const dayCell = sheet.getCell(row, 3);
      monthCell.value = monthStr;
      dayCell.value = dayStr;
      (monthCell as ExcelJS.Cell).numFmt = "@";
      (dayCell as ExcelJS.Cell).numFmt = "@";
      sheet.getCell(row, 4).value = weekday ?? "";
      (sheet.getCell(row, 4) as ExcelJS.Cell).alignment = { horizontal: "center", vertical: "middle" };

      const summaryList: string[] = [];
      let totalAmount = 0;

      for (let c = 0; c < menuOrder.length; c++) {
        const col = 5 + c;
        const menuName = menuOrder[c];
        const count = Number(counts[menuName]) || 0;
        const price = Number(prices[menuName]) || 0;
        const subTotal = price * count;
        const cell = sheet.getCell(row, col);
        cell.value = subTotal;
        (cell as ExcelJS.Cell).numFmt = '"￥"#,##0';
        (cell as ExcelJS.Cell).alignment = { horizontal: "center", vertical: "middle" };
        totalAmount += subTotal;
        if (count > 0) summaryList.push(`${menuName}\u3000${count}名`);
      }

      // 値引きの適用
      if (totalDiscount > 0) {
        totalAmount -= totalDiscount;
        summaryList.push(`（値引き：¥${totalDiscount.toLocaleString()}）`);
      }

      // 備考の追加
      if (remarks) {
        summaryList.push(`備考: ${remarks}`);
      }

      grandTotal += totalAmount;

      const priceCell = sheet.getCell(row, 11) as ExcelJS.Cell;
      priceCell.value = totalAmount;
      priceCell.numFmt = '"￥"#,##0';
      const centerAlign = { horizontal: "center" as const, vertical: "middle" as const };
      priceCell.alignment = centerAlign;
      priceCell.style = { ...priceCell.style, alignment: centerAlign };

      sheet.getCell(row, 13).value = summaryList.join("\n");
      (sheet.getCell(row, 13) as ExcelJS.Cell).alignment = { horizontal: "left", vertical: "top", wrapText: true };
    }

    // 14-15行: D14:G15（結合）に合計金額（中央揃え）
    const d14 = sheet.getCell("D14") as ExcelJS.Cell;
    d14.value = grandTotal;
    d14.numFmt = '"￥"#,##0';
    const centerAlign = { horizontal: "center" as const, vertical: "middle" as const };
    d14.alignment = centerAlign;
    d14.style = { ...d14.style, alignment: centerAlign };

    // 列幅: 備考(M,N)は狭く、データ行の高さを増やす
    const colWidths: Record<string, number> = {
      B: 6, C: 6, D: 6, E: 10, F: 10, G: 10, H: 10, I: 10, J: 10, K: 8, L: 8, M: 8, N: 8,
    };
    for (const [col, w] of Object.entries(colWidths)) {
      sheet.getColumn(col).width = w;
    }
    for (const row of dataRows) {
      sheet.getRow(row).height = 36;
      sheet.getRow(row + 1).height = 36;
    }

    // K11:O14 は会社名・住所・登録番号でテンプレートのまま。K36のみクリア
    try {
      sheet.getCell("K36").value = "";
    } catch {
      /* セルが存在しない場合は無視 */
    }

    const buffer = await workbook.xlsx.writeBuffer();
    return new NextResponse(buffer, {
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="order.xlsx"`,
      },
    });
  } catch (err: unknown) {
    console.error("Export error:", err);
    const message = err instanceof Error ? err.message : "Unknown error";
    return NextResponse.json({ error: "Excel export failed", details: message }, { status: 500 });
  }
}
