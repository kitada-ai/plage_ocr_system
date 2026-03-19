import { NextResponse } from "next/server";
import path from "path";
import fs from "fs";

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

export async function POST(req: Request) {
  try {
    const body = await req.json();
    const facilityName = body.facilityName;
    const dateStr = body.date || ""; 
    const [year, month, day] = dateStr.split("-");

    // ユーザー指定の固定テンプレート「テンプレ.xlsx」
    const templatePath = path.join(process.cwd(), "docs", "samples", "written", "テンプレ.xlsx");

    if (!fs.existsSync(templatePath)) {
      return NextResponse.json({ error: "テンプレート(テンプレ.xlsx)が見つかりません" }, { status: 500 });
    }

    try {
      const XlsxPopulate = require("xlsx-populate");
      const workbook = await XlsxPopulate.fromFileAsync(templatePath);
      
      // 最初のシート（Sheet1 または 申込書）を使用
      const sheet = workbook.sheet(0);

      // --- 1. 基本情報の書き出し (3行目) ---
      // I3: 施設名
      sheet.cell("I3").value(`施設名： ${facilityName || ""}`);
      // N3: 施術日
      const yearInt = parseInt(year || "0");
      const reiwa = yearInt > 2018 ? yearInt - 2018 : "  ";
      sheet.cell("N3").value(`令和 ${reiwa} 年 ${month || "  "} 月 ${day || "  "} 日`);

      // --- 2. メニューと料金の動的ヘッダー (13-14行目) ---
      const menuItems = body.menuItems || [];
      menuItems.forEach((item: any, i: number) => {
        if (i < 8) { // G列(7)から N列(14)付近までの枠
          const col = String.fromCharCode(71 + i); // G, H, I, J, K, L, M, N
          sheet.cell(`${col}13`).value(item.name || "");
          sheet.cell(`${col}14`).value(item.price || "");
        }
      });

      // --- 3. 実データの流し込み (17行目以降) ---
      const customers = body.customers || [];
      const rowDataStart = 17;
      
      customers.forEach((customer: CustomerData, idx: number) => {
        const rowNum = rowDataStart + idx;
        const row = sheet.row(rowNum);
        
        // C: 部屋番号, D: 氏名, F: 性別
        row.cell(3).value(customer.room || "");
        row.cell(4).value(customer.name || "");
        row.cell(6).value(customer.gender || "");

        // G列以降: メニュー〇付け
        menuItems.forEach((m: any, mIdx: number) => {
          if (mIdx < 8) {
            if (customer.selectedMenus && customer.selectedMenus.includes(m.name)) {
              row.cell(7 + mIdx).value("〇");
            }
          }
        });

        // O: 合計料金 はテンプレート側を維持するためノータッチ

        // P, Q, R: 第一〜第三希望
        const times = customer.preferredTimes || [];
        times.forEach((t: string, i: number) => { if (i < 3) row.cell(16 + i).value(t); });

        // S: ご案内, T: 実施
        if (customer.hasService) row.cell(20).value("✓");
        
        // U: 備考
        row.cell(21).value(customer.remarks || "");
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
