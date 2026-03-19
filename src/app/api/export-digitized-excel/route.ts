import { NextResponse } from "next/server";
import path from "path";
import fs from "fs";
import { exec } from "child_process";
import { promisify } from "util";
import os from "os";

const execAsync = promisify(exec);

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

    // テンプレートファイル
    const templatePath = path.join(
      process.cwd(),
      "docs",
      "samples",
      "sheets",
      "【★ﾄﾞｸﾀｰｻﾝｺﾞ守口様】訪問施術サービス （申込書・請求書）.xlsx"
    );

    if (!fs.existsSync(templatePath)) {
      return NextResponse.json(
        { error: "テンプレートファイルが見つかりません" },
        { status: 500 }
      );
    }

    // Pythonスクリプトのパス
    const scriptPath = path.join(process.cwd(), "scripts", "export_excel.py");

    // 一時ファイル用ディレクトリ
    const tempDir = os.tmpdir();
    const sessionId = Math.random().toString(36).substring(7);
    const payloadPath = path.join(tempDir, `payload_${sessionId}.json`);
    const outputPath = path.join(tempDir, `output_${sessionId}.xlsx`);

    // ペイロードを一時ファイルに保存
    fs.writeFileSync(payloadPath, JSON.stringify(body, null, 2), "utf8");

    try {
      // Pythonスクリプトを実行
      // Windows環境なので python または python.exe
      const pythonCmd = process.platform === "win32" ? "python" : "python3";
      await execAsync(`"${pythonCmd}" "${scriptPath}" "${payloadPath}" "${templatePath}" "${outputPath}"`);

      // 生成されたファイルを読み込み
      const buffer = fs.readFileSync(outputPath);

      // 一時ファイルを削除
      try {
        fs.unlinkSync(payloadPath);
        fs.unlinkSync(outputPath);
      } catch (e) {
        console.warn("一時ファイルの削除に失敗しました:", e);
      }

      const formatDate = `${year || ""}${month || ""}${day || ""}`;
      const fileName = `${facilityName || "申込書"}_${formatDate || "清書"}.xlsx`
        .replace(/[/\\?%*:|"<>]/g, "_");

      return new NextResponse(buffer, {
        status: 200,
        headers: {
          "Content-Disposition": `attachment; filename="${encodeURIComponent(fileName)}"`,
          "Content-Type":
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        },
      });
    } catch (pythonError: any) {
      console.error("Python Execution Error:", pythonError);
      return NextResponse.json(
        { error: "PythonでのExcel生成に失敗しました", details: pythonError.stderr || pythonError.message },
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

