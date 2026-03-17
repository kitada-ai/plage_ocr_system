import { NextResponse } from "next/server";
import ExcelJS from "exceljs";

export async function POST(req: Request) {
  try {
    const { headers, rows, prices } = await req.json();

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("手書きデジタイズ結果");

    // 1. Setup Headers
    worksheet.columns = headers.map((h: string) => ({
      header: h,
      key: h,
      width: h === "氏名" ? 25 : 12
    }));

    // Add Pricing row (optional hint)
    const pricingRow: any = { "氏名": "マスタ単価" };
    headers.forEach((h: string) => {
      if (h !== "氏名" && h !== "施術実施") {
        pricingRow[h] = prices[h] || 0;
      }
    });
    worksheet.addRow(pricingRow);
    worksheet.getRow(2).font = { italic: true, color: { argb: "FF888888" } };

    // 2. Add Data Rows
    rows.forEach((row: any) => {
      const rowData: any = {};
      headers.forEach((h: string) => {
        rowData[h] = row[h] || "";
      });
      worksheet.addRow(rowData);
    });

    // Styling
    worksheet.getRow(1).font = { bold: true };
    worksheet.getRow(1).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFF2F2F2" }
    };

    // Buffer to response
    const buffer = await workbook.xlsx.writeBuffer();

    return new NextResponse(buffer, {
      status: 200,
      headers: {
        "Content-Disposition": 'attachment; filename="digitized_data.xlsx"',
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      },
    });

  } catch (error: any) {
    console.error("Excel Export Error:", error);
    return NextResponse.json({ error: "Excel生成に失敗しました", details: error.message }, { status: 500 });
  }
}
