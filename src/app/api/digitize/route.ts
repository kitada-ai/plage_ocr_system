import { NextResponse } from "next/server";

export async function POST(req: Request) {
  try {
    const formData = await req.formData();
    const file = formData.get("file") as File;
    if (!file) {
      return NextResponse.json({ error: "ファイルがありません" }, { status: 400 });
    }

    const azureEndpoint = process.env.AZURE_DI_ENDPOINT;
    const azureKey = process.env.AZURE_DI_KEY;
    const geminiKey = process.env.GEMINI_API_KEY;

    if (!azureEndpoint || !azureKey || !geminiKey) {
      return NextResponse.json({ error: "APIキーが設定されていません" }, { status: 500 });
    }

    // --- 1. Azure AI Document Intelligence (Layout) ---
    const arrayBuffer = await file.arrayBuffer();
    const analyzeUrl = `${azureEndpoint.replace(/\/$/, "")}/formrecognizer/documentModels/prebuilt-layout:analyze?api-version=2023-07-31`;
    
    const azureRes = await fetch(analyzeUrl, {
      method: "POST",
      headers: {
        "Content-Type": file.type || "application/octet-stream",
        "Ocp-Apim-Subscription-Key": azureKey,
      },
      body: arrayBuffer,
    });

    if (!azureRes.ok) {
      const errorText = await azureRes.text();
      return NextResponse.json({ error: "Azure解析失敗", details: errorText }, { status: azureRes.status });
    }

    const operationLocation = azureRes.headers.get("operation-location");
    if (!operationLocation) {
      return NextResponse.json({ error: "Azure operation-locationが見つかりません" }, { status: 500 });
    }

    // Polling Azure result
    let azureResult: any = null;
    for (let i = 0; i < 15; i++) {
      await new Promise(r => setTimeout(r, 1000));
      const pollRes = await fetch(operationLocation, {
        headers: { "Ocp-Apim-Subscription-Key": azureKey },
      });
      const pollJson = await pollRes.json();
      if (pollJson.status === "succeeded") {
        azureResult = pollJson.analyzeResult;
        break;
      }
      if (pollJson.status === "failed") throw new Error("Azure解析失敗");
    }

    if (!azureResult) return NextResponse.json({ error: "Azure解析タイムアウト" }, { status: 500 });

    // --- 2. Extract Table Metadata for Gemini ---
    const table = azureResult.tables?.[0]; // Assume first table for simplicity for now
    if (!table) return NextResponse.json({ error: "テーブルが見つかりませんでした" });

    // Filter interesting info for Gemini (Reducing token count)
    const tableSummary = {
      rowCount: table.rowCount,
      columnCount: table.columnCount,
      cells: table.cells.map((c: any) => ({
        rowIndex: c.rowIndex,
        columnIndex: c.columnIndex,
        content: c.content,
        // We could provide coordinates but Gemini Vision can see the image anyway.
        // Providing hint about layout is enough.
      }))
    };

    // --- 3. Google Gemini (VLM Semantic Interpretation) ---
    // Gemini 1.5 Flash supports image + text
    const base64Image = Buffer.from(arrayBuffer).toString('base64');
    
    const geminiPrompt = `
      添付された理・美容施設の「施術記録表」の画像と、Azure OCRによって抽出された表構造（JSON）を元に、データを正確にデジタル化してください。
      
      【ミッション】
      1. 手書きの記号（〇、チェック、斜線など）を読み取り、実施されたメニューを特定してください。
      2. 「氏名」を正確に読み取ってください。
      3. メニュー名（ヘッダー）を、一般的な名称に正規化してください。
      4. 各行がAzure解析結果のどの「rowIndex」に対応するかを記録してください。
      
      【Azure解析結果のヒント】
      ${JSON.stringify(tableSummary)}
      
      【回答形式】
      以下のJSON形式のみで回答してください：
      {
        "headers": ["氏名", "カット", "カラー", ...],
        "rows": [
          { 
            "氏名": "山田 太郎", 
            "カット": "〇", 
            "カラー": "×", 
            ... ,
            "azureRowIndex": 5
          },
          ...
        ],
        "prices": {
          "カット": 1500,
          "カラー": 3500
        }
      }
    `;

    const geminiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${geminiKey}`;
    
    const geminiRes = await fetch(geminiUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        contents: [{
          parts: [
            { text: geminiPrompt },
            {
              inline_data: {
                mime_type: file.type,
                data: base64Image
              }
            }
          ]
        }],
        generationConfig: {
          response_mime_type: "application/json",
        }
      })
    });

    if (!geminiRes.ok) {
      const errorText = await geminiRes.text();
      return NextResponse.json({ error: "Gemini解析失敗", details: errorText }, { status: geminiRes.status });
    }

    const geminiData = await geminiRes.json();
    const resultText = geminiData.candidates[0].content.parts[0].text;
    const finalResult = JSON.parse(resultText);

    return NextResponse.json({
      ...finalResult,
      azureTableData: table // Return raw table for coordinate mapping
    });

  } catch (error: any) {
    console.error(error);
    return NextResponse.json({ error: "解析プロセスでエラー", details: error.message }, { status: 500 });
  }
}
