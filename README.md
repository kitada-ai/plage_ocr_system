# 請求書作成支援システム

理容・美容施設の施術記録表を画像から自動で読み取り、Excelの請求書として出力するシステムです。

## 主な機能

### 1. 画像のOCR処理
- Azure AI Document Intelligenceを使用して施術記録表を読み取り
- PDF・画像ファイル（複数枚）に対応
- 画像の回転機能付き

### 2. 記号の自動識別
- Custom Visionを使って記号（〇、×、チェック、斜線など）を自動判定
- 氏名、カット、カラー、パーマ、ヘアーマニキュア、ベットカット、顔そり、シャンプー、施術実施などの項目を動的に検出

### 3. データの編集機能
- 識別結果をUI上で確認・修正可能
- セルをクリックして〇/×を切り替え
- 列全体の一括切り替え機能
- 施術実施列の一括編集モーダル（全体/画像単位/ページ単位）
- メニュー列の追加・削除
- 人（行）の追加・削除

### 4. 文書情報の自動抽出
- Gemini APIで施設名、施術日、曜日などを自動抽出
- 令和年号に自動変換

### 5. 集計とExcel出力
- メニュー別に施術人数を集計
- 単価×人数で金額を自動計算
- テンプレートベースでExcel請求書を生成
- 会社情報と振込先情報を自動配置

### 6. その他の便利機能
- テーブルのズーム表示（50%〜200%）
- 画像プレビューの拡大表示
- デバッグモード（切り取り画像の確認用）
- 処理進捗の表示

## 技術スタック

- **フロントエンド**: Next.js 16, React 19, TypeScript
- **OCR**: Azure AI Document Intelligence (layoutモデル)
- **記号識別**: Azure Custom Vision
- **AI**: Google Gemini API（文書情報抽出）
- **Excel処理**: ExcelJS
- **PDF処理**: pdf.js

## セットアップ

**別のパソコンでクローンして動かす手順**は [docs/クローンとセットアップ.md](docs/クローンとセットアップ.md) を参照してください。

1. 依存関係のインストール:
```bash
npm install
```

2. 環境変数の設定（`.env.local`ファイルを作成）:
```
# Azure Document Intelligence（OCR）
AZURE_DI_ENDPOINT=あなたのエンドポイント
AZURE_DI_KEY=あなたのキー

# Azure Custom Vision
NEXT_PUBLIC_CUSTOM_VISION_KEY=あなたのキー
NEXT_PUBLIC_CUSTOM_VISION_ENDPOINT=あなたのエンドポイント
NEXT_PUBLIC_CUSTOM_VISION_PROJECT_ID=あなたのプロジェクトID
NEXT_PUBLIC_CUSTOM_VISION_ITERATION_ID=あなたのイテレーションID

# Google Gemini API
GEMINI_API_KEY=あなたのAPIキー
```

3. 開発サーバーの起動:
```bash
npm run dev
```

## 使い方

1. **画像をアップロード**: 複数のPDFや画像ファイルを選択
2. **画像の回転**: 必要に応じて画像を90度ずつ回転
3. **解析開始**: 「アップロード & 解析」ボタンをクリック
4. **結果の確認と修正**: OCR結果を確認し、必要に応じて修正
5. **集計確定**: 「集計確定」ボタンで施術実施のデータを集計
6. **Excel出力**: 「Excel出力」ボタンで請求書をダウンロード

## プロジェクト構造

```
ocr-main/
├── src/                    # アプリケーションソース
│   ├── app/                # Next.js App Router
│   │   ├── page.tsx        # メインUI
│   │   ├── login/page.tsx  # ログイン
│   │   └── api/            # API Routes（analyze, gemini, export, auth 等）
│   └── proxy.ts            # 認証ガード（Next.js 16）
├── config/                 # ツール設定
│   └── eslint.config.mjs
├── public/
│   └── templates/         # Excel請求書テンプレート
├── docs/                   # ドキュメント・サンプル・dev-log
├── scripts/
│   └── main.py            # スタンドアロンOCRスクリプト（開発用）
├── next.config.ts         # Next.js 設定
├── tsconfig.json          # TypeScript 設定
├── postcss.config.mjs     # PostCSS 設定
├── package.json
├── .env.example            # 環境変数の雛形
└── README.md
```

スタンドアロンOCR（`scripts/main.py`）を実行する場合、プロジェクトルートで環境変数を設定したうえで `python scripts/main.py` を実行してください。入力ファイルは `docs/samples/images/sample.jpg` を参照します。

## リポジトリに含まないもの（プッシュしない）

以下のファイル・フォルダは `.gitignore` で除外され、Git にはコミットしません。ローカルまたは CI/本番環境で別途用意します。

| 対象 | 説明 |
|------|------|
| `.env.local` / `env` | 環境変数（APIキー等）。`.env.example` をコピーして作成 |
| `node_modules/` | 依存パッケージ。`npm install` で生成 |
| `.next/` / `out/` / `build/` | ビルド成果物 |
| `.vercel/` | Vercel のプロジェクト紐づけ情報 |
| `logs/` | スタンドアロンOCRのログ出力先 |
| `.claude` / `.cursor` | エディタ・AI のローカル設定 |
| `~$*` | Excel の一時ロックファイル |
| `*.tsbuildinfo` / `next-env.d.ts` | 自動生成ファイル |

## 既知の制限事項

- Azure Document Intelligence Free Tier (F0) を使用している場合、クォータ制限に注意
- レート制限対策として、画像処理間に1秒の待機時間を設定
- レート制限エラー時は自動的に最大5回までリトライ
- メニュー数は最大12個まで（Excel出力時）
