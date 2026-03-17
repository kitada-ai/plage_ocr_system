# API料金見積もり - 月10枚の書類処理

**作成日**: 2026年2月2日  
**対象**: 月10枚の書類を処理する場合のAPI料金見積もり

---

## 📊 システム構成とAPI使用状況

このシステムでは以下の3つのAPIを使用しています：

1. **Azure Document Intelligence** - OCR処理（文書解析）
2. **Google Gemini API** - テキストからの情報抽出
3. **Azure Custom Vision** - テーブルセルの画像分類

---

## 💰 各APIの料金体系

### 1. Azure Document Intelligence (prebuilt-layout)

**料金**:
- **無料プラン**: 月500ページまで無料
- **従量課金**: 1,000ページあたり $1.50（約¥225、為替レート1ドル=¥150想定）

**使用状況（月10枚想定）**:
- 1枚の書類 = 1ページとして処理
- 月10枚 = **10ページ**
- **料金**: **¥0（無料プラン内）**

---

### 2. Google Gemini API (gemini-2.5-flash)

**料金**:
- **入力トークン**: ¥45/100万トークン
- **出力トークン**: ¥375/100万トークン

**使用状況（月10枚想定）**:
- 1枚の書類につき1回のAPI呼び出し
- 月10枚 = **10回のAPI呼び出し**

**トークン使用量の推定**:
- **入力トークン**: プロンプト（約500トークン）+ 抽出対象テキスト（平均2,000トークン）= **約2,500トークン/回**
  - 月10回 = **25,000トークン** = **0.025百万トークン**
  - 料金: 0.025 × ¥45 = **¥1.125**

- **出力トークン**: JSON形式のレスポンス（約50トークン/回）
  - 月10回 = **500トークン** = **0.0005百万トークン**
  - 料金: 0.0005 × ¥375 = **¥0.188**

- **合計**: **約¥1.31/月**

---

### 3. Azure Custom Vision (Prediction API)

**料金**:
- 従量課金制（トランザクション数ベース）
- 詳細な料金は公式ページで確認が必要
- 一般的に1回あたり **$0.001～$0.01**（約¥0.15～¥1.5）程度と推定

**使用状況（月10枚想定）**:
- テーブルの各セルごとに画像分類APIを呼び出し
- 1枚の書類あたりのセル数は可変（テーブル構造による）

**推定**:
- 1枚の書類に平均 **50セル** があると仮定
- 月10枚 = **500セル** = **500回のAPI呼び出し**
- 1回あたり **¥0.5** と仮定
- **料金**: 500 × ¥0.5 = **¥250/月**

**注意**: Custom Visionの実際の料金は、プロジェクトの設定やリージョンによって異なります。正確な料金は[Azure Custom Vision価格ページ](https://azure.microsoft.com/ja-jp/pricing/details/cognitive-services/custom-vision-service/)で確認してください。

---

## 📈 月間料金見積もり（月10枚）

| API | 使用量 | 単価 | 月間料金 |
|-----|--------|------|----------|
| **Azure Document Intelligence** | 10ページ | 無料（500ページまで） | **¥0** |
| **Google Gemini API** | 10回（約25,000入力 + 500出力トークン） | 入力¥45/100万、出力¥375/100万 | **¥1.31** |
| **Azure Custom Vision** | 約500回（推定） | 約¥0.5/回（推定） | **¥250** |
| **合計** | - | - | **約¥251/月** |

---

## 📝 詳細な内訳

### 最安ケース（Custom Vision使用量が少ない場合）
- Azure Document Intelligence: ¥0
- Google Gemini API: ¥1.31
- Azure Custom Vision: ¥50（100回/月と仮定）
- **合計: 約¥51/月**

### 高使用ケース（Custom Vision使用量が多い場合）
- Azure Document Intelligence: ¥0
- Google Gemini API: ¥1.31
- Azure Custom Vision: ¥500（1,000回/月と仮定）
- **合計: 約¥501/月**

---

## ⚠️ 注意事項

1. **Azure Document Intelligence**
   - 無料プラン（月500ページ）を超える場合は、1,000ページあたり約¥225の追加料金が発生します
   - 月10枚の使用量では無料プラン内です

2. **Google Gemini API**
   - トークン使用量は実際の文書の長さによって変動します
   - 長文の書類が多い場合は、入力トークンが増加する可能性があります

3. **Azure Custom Vision**
   - テーブルのセル数によってAPI呼び出し回数が大きく変動します
   - 正確な料金は、実際の使用量を確認してから計算してください
   - 無料プランや割引プランが適用される場合があります

4. **為替レート**
   - ドル建ての料金は為替レートによって変動します
   - 上記の見積もりは1ドル=¥150を想定しています

---

## 💡 コスト削減のヒント

1. **Azure Document Intelligence**: 月500ページまで無料なので、現状の使用量では追加コストは発生しません

2. **Google Gemini API**: 
   - プロンプトを最適化してトークン数を削減
   - 必要最小限のテキストのみを送信

3. **Azure Custom Vision**:
   - 不要なセルの分類をスキップ
   - バッチ処理で効率化
   - 無料プランや割引プランの確認

---

## 📞 正確な料金確認方法

1. **Azure Document Intelligence**: 
   - [Azure料金計算ツール](https://azure.microsoft.com/ja-jp/pricing/calculator/)
   - [Azure Document Intelligence価格ページ](https://azure.microsoft.com/ja-jp/pricing/details/ai-document-intelligence/)

2. **Google Gemini API**:
   - [Google AI Studio](https://aistudio.google.com/)で使用量を確認
   - [Gemini API価格ページ](https://ai.google.dev/pricing)

3. **Azure Custom Vision**:
   - [Azure Custom Vision価格ページ](https://azure.microsoft.com/ja-jp/pricing/details/cognitive-services/custom-vision-service/)
   - Azure Portalで実際の使用量を確認

---

## 📊 まとめ

**月10枚の書類処理における推定API料金: 約¥250～¥500/月**

- 主なコスト要因: Azure Custom Vision（テーブルセルの画像分類）
- 無料プランの活用: Azure Document Intelligenceは無料プラン内
- 変動要因: Custom Visionの使用回数（テーブルのセル数に依存）

実際の使用量に基づいて、より正確な見積もりを算出することをお勧めします。
