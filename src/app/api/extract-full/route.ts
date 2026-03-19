import { NextResponse } from "next/server";

/**
 * OCRテキストからフル情報を一括抽出するAPI
 * Gemini APIを使用して施設名・日付・メニュー・顧客データ・希望時間等を抽出
 */
export async function POST(req: Request) {
  try {
    const { text } = await req.json();

    const apiKey = process.env.GEMINI_API_KEY || "";
    const modelName = process.env.GEMINI_MODEL_NAME || "gemini-2.5-flash";
    const endpoint =
      process.env.GEMINI_ENDPOINT ||
      `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent`;

    if (!apiKey) {
      return NextResponse.json(
        { error: "APIキーが設定されていません" },
        { status: 500 }
      );
    }

    const now = new Date();
    const currentYear = now.getFullYear();
    const currentMonth = now.getMonth() + 1;

    const prompt = `
以下のテキストは、訪問理美容サービスの申込書をOCRで読み取ったものです。
以下のJSON形式で全ての情報を抽出してください。JSON以外の文字は含めないでください。

{
  "facilityName": "施設名（例：ドクターサンゴ守口）",
  "year": "年（4桁）",
  "month": "月",
  "day": "日",
  "dayOfWeek": "曜日（月火水木金土日のいずれか1文字）",
  "menuItems": [
    { "name": "メニュー名", "price": 数値 }
  ],
  "customers": [
    {
      "no": 番号,
      "room": "部屋番号",
      "name": "氏名",
      "gender": "男 or 女 or 不明",
      "selectedMenus": ["選択されているメニュー名の配列"],
      "preferredTimes": ["第一希望の時間", "第二希望の時間", "第三希望の時間"],
      "hasService": true/false,
      "isGuided": "案内有無 (有/無/空欄など)",
      "isAdditionalMenuAllowed": "追加メニュー可否 (可/否/空欄など)",
      "isCustomOrder": "オーダーメイドの内容 (内容があれば文字列、なければ空欄)",
      "remarks": "備考"
    }
  ]
}

【抽出ルール】
1. メニュー名と料金はヘッダー行から抽出してください。〇がついているものが選択されたメニューです。
2. 顧客は「記入例」「山田 太郎」を除外してください。
3. 年の記載がない場合：現在は ${currentYear}年${currentMonth}月です。テキストの月が${currentMonth}より大きければ${currentYear - 1}年、それ以外は${currentYear}年。
4. 曜日がない場合は年月日から計算してください。
5. 施設名のフロア情報（2F、3階等）は除去してください。
6. 「施術開始時間の希望」の第一〜第三希望を全て抽出してください。ない場合は空文字列。
7. 「施術実施有無」に✓や〇がある場合はhasServiceをtrueにしてください。
8. 空の行（氏名が空）はスキップしてください。
9. メニューに〇がついているかの判断：「〇」「○」「✓」「チェック」「selected」等を〇と判断してください。

テキスト:
${text}
    `;

    let requestUrl = endpoint;
    if (
      !process.env.GEMINI_ENDPOINT &&
      !endpoint.includes(modelName)
    ) {
      requestUrl = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent`;
    } else if (
      process.env.GEMINI_ENDPOINT &&
      !endpoint.includes(modelName) &&
      endpoint.includes(":generateContent")
    ) {
      requestUrl = endpoint.replace(
        /\/models\/[^:]+/,
        `/models/${modelName}`
      );
    }

    const requestBody: Record<string, unknown> = {
      contents: [{ parts: [{ text: prompt }] }],
    };

    if (requestUrl.includes("/v1beta/")) {
      requestBody.generationConfig = {
        responseMimeType: "application/json",
        responseSchema: {
          type: "object",
          properties: {
            facilityName: { type: "string" },
            year: { type: "string" },
            month: { type: "string" },
            day: { type: "string" },
            dayOfWeek: { type: "string" },
            menuItems: {
              type: "array",
              items: {
                type: "object",
                properties: {
                  name: { type: "string" },
                  price: { type: "number" },
                },
                required: ["name", "price"],
              },
            },
            customers: {
              type: "array",
              items: {
                type: "object",
                properties: {
                  no: { type: "number" },
                  room: { type: "string" },
                  name: { type: "string" },
                  gender: { type: "string" },
                  selectedMenus: {
                    type: "array",
                    items: { type: "string" },
                  },
                  preferredTimes: {
                    type: "array",
                    items: { type: "string" },
                  },
                  hasService: { type: "boolean" },
                  isGuided: { type: "string" },
                  isAdditionalMenuAllowed: { type: "string" },
                  isCustomOrder: { type: "string" },
                  remarks: { type: "string" },
                },
                required: ["name"],
              },
            },
          },
          required: [
            "facilityName",
            "year",
            "month",
            "day",
            "dayOfWeek",
            "menuItems",
            "customers",
          ],
        },
      };
    } else {
      requestBody.generationConfig = { temperature: 0.1 };
    }

    console.log(`extract-full API呼び出し: ${requestUrl}`);

    const response = await fetch(requestUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-goog-api-key": apiKey,
      },
      body: JSON.stringify(requestBody),
    });

    if (!response.ok) {
      const errorText = await response.text();
      console.error("extract-full API Error:", response.status, errorText);
      return NextResponse.json(
        { error: `API呼び出しに失敗しました: ${response.status}` },
        { status: response.status }
      );
    }

    const result = await response.json();

    let extractedData;
    if (
      result.candidates?.[0]?.content
    ) {
      let content = result.candidates[0].content.parts[0].text;
      content = content
        .replace(/^```json\s*/i, "")
        .replace(/\s*```$/, "")
        .trim();
      extractedData = JSON.parse(content);
    } else if (result.text) {
      let text = result.text;
      text = text
        .replace(/^```json\s*/i, "")
        .replace(/\s*```$/, "")
        .trim();
      extractedData = JSON.parse(text);
    } else {
      extractedData = result;
    }

    return NextResponse.json(extractedData);
  } catch (error) {
    console.error("extract-full API Error:", error);
    return NextResponse.json(
      {
        error: "抽出に失敗しました",
        details: error instanceof Error ? error.message : String(error),
      },
      { status: 500 }
    );
  }
}
