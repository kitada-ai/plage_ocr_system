import { NextResponse } from "next/server";

export async function POST(req: Request) {
  try {
    const { menuCandidates } = await req.json();

    if (!menuCandidates || !Array.isArray(menuCandidates)) {
      return NextResponse.json(
        { error: "menuCandidates配列が必要です" },
        { status: 400 }
      );
    }

    const apiKey = process.env.GEMINI_API_KEY || "";
    const modelName = process.env.GEMINI_MODEL_NAME || "gemini-2.5-flash";

    if (!apiKey) {
      return NextResponse.json({ error: "APIキーが設定されていません" }, { status: 500 });
    }

    const prompt = `
以下のリストは美容院・理容室のサービスメニューの候補です。
各項目が実際の施術メニュー（カット、パーマ、カラーなど）であるかを判定してください。

【判定基準】
✓ 施術メニューである: カット、パーマ、カラー、シャンプー、顔そり、顔剃り、ベッドカット、ペットカット、ヘアーマニキュア、トリートメント、ヘッドスパなど具体的なサービス名
✗ 施術メニューではない: 時間、料金、金額、合計、メニュー（単体）、数字のみ、記号のみ、日付、曜日、時刻表記（9時、30分、9:00、10時30分など）、施術開始時間、希望、ご案内、案内

メニュー候補:
${menuCandidates.map((m: string, i: number) => `${i + 1}. ${m}`).join('\n')}

以下のJSON形式で回答してください（JSON以外の文字は含めないでください）:
{
  "validMenus": ["実際の施術メニューであるもののリスト"],
  "invalidMenus": ["施術メニューではないもののリスト"]
}
`;

    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent`;

    console.log(`🤖 Menu Validation API呼び出し: ${endpoint}`);

    const requestBody = {
      contents: [{
        parts: [{
          text: prompt
        }]
      }],
      generationConfig: {
        responseMimeType: "application/json",
        responseSchema: {
          type: "object",
          properties: {
            validMenus: {
              type: "array",
              items: { type: "string" },
              description: "実際の施術メニューであるもののリスト"
            },
            invalidMenus: {
              type: "array",
              items: { type: "string" },
              description: "施術メニューではないもののリスト"
            }
          },
          required: ["validMenus", "invalidMenus"]
        }
      }
    };

    const response = await fetch(endpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-goog-api-key": apiKey,
      },
      body: JSON.stringify(requestBody),
    });

    if (!response.ok) {
      const errorText = await response.text();
      console.error("Gemini API Error:", response.status, errorText);
      return NextResponse.json(
        { error: `API呼び出しに失敗しました: ${response.status}` },
        { status: response.status }
      );
    }

    const result = await response.json();

    let extractedData;
    if (result.candidates && result.candidates[0] && result.candidates[0].content) {
      let content = result.candidates[0].content.parts[0].text;

      // マークダウンコードブロックを除去
      content = content.replace(/^```json\s*/i, '').replace(/\s*```$/, '').trim();

      extractedData = JSON.parse(content);
    } else {
      extractedData = result;
    }

    console.log("✅ Menu Validation Result:", extractedData);

    return NextResponse.json(extractedData);
  } catch (error) {
    console.error("Menu Validation Error:", error);
    return NextResponse.json(
      {
        error: "メニュー検証に失敗しました",
        details: error instanceof Error ? error.message : String(error)
      },
      { status: 500 }
    );
  }
}
