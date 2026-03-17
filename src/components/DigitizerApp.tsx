"use client";

import { useState, useEffect, useRef } from "react";

// --- Types ---
type TableCell = {
  rowIndex: number;
  columnIndex: number;
  content?: string;
  columnSpan?: number;
  rowSpan?: number;
  boundingRegions?: { polygon: number[] }[];
};

interface BoundingRegion {
  pageNumber: number;
  polygon: number[];
}

interface Table {
  rowCount: number;
  columnCount: number;
  cells: TableCell[];
  boundingRegions?: BoundingRegion[];
}

type DisplayRow = {
  rowIndex: number;
  /** 行の一意ID（トグル時に1行だけ更新するために使用。未設定時は rowIndex で識別） */
  rowId?: number;
  columns: string[];
  results: (string | null)[];
  sourceImageIndex?: number;
  sourceImageName?: string;
  groupId?: number;
  polygon?: number[];
  cellPolygons?: Record<number, number[]>;
  discount?: number;
  remarks?: string;
};

type DocInfo = {
  facilityName: string;
  year: string;
  month: string;
  day: string;
  dayOfWeek: string;
};

type PageImage = {
  blob: Blob;
  imageUrl: string;
  pageNumber: number;
  width: number;
  height: number;
  rotation: number; // 0, 90, 180, 270
  fileName: string;
};

// --- Constants ---
const TARGET_COLUMNS = [
  "氏名",
  "カット",
  "カラー",
  "パーマ",
  "ヘアーマニキュア",
  "ベットカット",
  "ペットカット",
  "顔そり",
  "シャンプー",
  "施術実施",
];

// デフォルトの価格マップ（メニュー名に価格が含まれていない場合に使用）
const DEFAULT_MENU_PRICES: Record<string, number> = {
  "カット": 1800,
  "カラー": 3800,
  "カラー（白髪染め）": 3800,
  "パーマ": 3800,
  "ヘアーマニキュア": 2500,
  "ヘアマニキュア": 2500,
  "ベットカット": 1800,
  "ベッドカット": 1800,
  "ペットカット": 1800,
  "顔そり": 500,
  "顔剃り": 500,
  "シャンプー": 800,
  "パッチテスト": 0,
  "白髪染め": 3800,
  "トリートメント": 1500,
  "ヘッドスパ": 1500,
};

/** ヘッダー名から価格を取得（¥表記付きメニュー名やバリエーションにも対応） */
function getPriceForHeader(header: string, prices: Record<string, number> | undefined): number | undefined {
  if (!prices) return undefined;
  if (prices[header] !== undefined) return prices[header];
  const baseName = header.replace(/[\s\u3000]*[¥￥][\d,]+.*$/, "").trim();
  if (baseName && prices[baseName] !== undefined) return prices[baseName];
  if (/カラー.*白髪染め|白髪染め/.test(header)) {
    if (prices["カラー（白髪染め）"] !== undefined) return prices["カラー（白髪染め）"];
    if (prices["カラー"] !== undefined) return prices["カラー"];
  }
  if (header === "顔剃り" && prices["顔そり"] !== undefined) return prices["顔そり"];
  return undefined;
}

const CUSTOM_VISION_API_KEY = process.env.NEXT_PUBLIC_CUSTOM_VISION_KEY || "";
const CUSTOM_VISION_ENDPOINT = process.env.NEXT_PUBLIC_CUSTOM_VISION_ENDPOINT || "";
const PROJECT_ID = process.env.NEXT_PUBLIC_CUSTOM_VISION_PROJECT_ID || "";
const ITERATION_ID = process.env.NEXT_PUBLIC_CUSTOM_VISION_ITERATION_ID || "";

// --- ヘルパー関数: 複数テーブルの解析結果を連結 ---
function mergeDisplayResults(results: { displayRows: DisplayRow[], indices: number[], headers: string[], xPositions?: number[], prices?: Record<string, number> }[]) {
  console.log(`  📊 mergeDisplayResults: ${results.length}個の結果を処理`);
  if (results.length === 0) return null;

  // メニューヘッダーとX座標情報を収集
  const menuHeaderMap = new Map<string, number>(); // header -> X座標
  const globalMenuPrices: Record<string, number> = {}; // 価格情報をマージ

  const normalizeMenuHeader = (h: string) => (h === "顔剃り" ? "顔そり" : h);

  // メニューではない列（男女・性別・selected/unselected・例示名等）はマージ時に除外
  const isNonMenuHeader = (h: string): boolean => {
    const t = h.trim().toLowerCase();
    if (h === "氏名" || h === "施術実施") return true;
    if (t === "男女" || t === "性別") return true;
    if (t.includes("unselected") || t.includes("selected") || t === ":unselected:" || t === ":selected:") return true;
    if (h === "太郎" || h === "山田") return true;  // 例示行の名前がカラムに入らないように
    if (/^[※]/.test(h)) return true;
    return false;
  };

  results.forEach(({ headers, xPositions, prices }) => {
    headers.forEach((header, index) => {
      if (header === "氏名" || header === "施術実施") return;
      if (isNonMenuHeader(header)) return;

      const normalized = normalizeMenuHeader(header);
      const xPos = xPositions?.[index] ?? index * 100;
      if (!menuHeaderMap.has(normalized) || xPos < (menuHeaderMap.get(normalized) ?? Infinity)) {
        menuHeaderMap.set(normalized, xPos);
      }

      if (prices && prices[header] && !globalMenuPrices[normalized]) {
        globalMenuPrices[normalized] = prices[header];
      }
    });
  });

  // メニューヘッダーをX座標でソート（非メニューは除外）
  let sortedMenuHeaders = Array.from(menuHeaderMap.entries())
    .filter(([header]) => !isNonMenuHeader(header))
    .sort((a, b) => a[1] - b[1])
    .map(([header]) => header);

  // カットを強制的に最初に移動
  const cutIndex = sortedMenuHeaders.indexOf("カット");
  if (cutIndex > 0) {
    console.log(`  🔧 カットを先頭に移動: ${cutIndex}番目 → 0番目`);
    sortedMenuHeaders = ["カット", ...sortedMenuHeaders.filter(h => h !== "カット")];
  }

  console.log(`  📍 X座標ソート後のグローバルヘッダー: ${sortedMenuHeaders.join(', ')}`);
  if (Object.keys(globalMenuPrices).length > 0) {
    console.log(`  💰 マージされた価格情報:`, globalMenuPrices);
  }

  const globalHeaders = ["氏名", ...sortedMenuHeaders, "施術実施"];
  const globalIndices = globalHeaders.map((_, idx) => idx);

  const normalizeRows = (result: { displayRows: DisplayRow[], headers: string[] }) => {
    const localToGlobal = new Map<number, number>();
    result.headers.forEach((header, localIndex) => {
      const globalIndex = globalHeaders.indexOf(normalizeMenuHeader(header));
      if (globalIndex !== -1) localToGlobal.set(localIndex, globalIndex);
    });

    return result.displayRows.map((row) => {
      const isHeaderRow = row.results.every((res) => res === null);
      const newColumns = Array(globalHeaders.length).fill("");
      const newResults = Array(globalHeaders.length).fill(null) as (string | null)[];

      if (isHeaderRow) {
        globalHeaders.forEach((header, index) => {
          newColumns[index] = header;
        });
      } else {
        localToGlobal.forEach((globalIndex, localIndex) => {
          newColumns[globalIndex] = row.columns[localIndex] ?? "";
        });
      }

      localToGlobal.forEach((globalIndex, localIndex) => {
        newResults[globalIndex] = row.results[localIndex] ?? null;
      });

      return { ...row, columns: newColumns, results: newResults };
    });
  };

  const baseResult = results[0];
  const baseDisplayRows = normalizeRows(baseResult).filter((row, index) => {
    const isHeaderRow = row.results.every((res) => res === null);
    const isYamadaRow = row.columns[0] === "山田 太郎";
    // 最初のヘッダー行は残す、それ以外のヘッダーと山田行はスキップ
    if (index === 0 && isHeaderRow) return true;
    return !isHeaderRow && !isYamadaRow;
  });
  console.log(`    結果[0]: ${baseDisplayRows.length}行 (ベース、ヘッダー1行+データ行のみ)`);

  const mergedRows = [...baseDisplayRows];

  for (let i = 1; i < results.length; i++) {
    const current = results[i];
    console.log(`    結果[${i}]: ${current.displayRows.length}行 (処理前)`);

    const normalizedRows = normalizeRows(current).filter((row) => {
      const isHeaderRow = row.results.every((res) => res === null);
      const isYamadaRow = row.columns[0] === "山田 太郎";
      const shouldSkip = isHeaderRow || isYamadaRow;

      if (shouldSkip) {
        console.log(`      スキップ: ${row.columns[0]} (ヘッダー=${isHeaderRow}, 山田=${isYamadaRow})`);
      }

      return !shouldSkip;
    });

    console.log(`    → ${normalizedRows.length}行を追加 (フィルタ後)`);
    mergedRows.push(...normalizedRows);
  }

  console.log(`  ✅ マージ後の合計: ${mergedRows.length}行`);
  return {
    displayRows: mergedRows,
    indices: globalIndices,
    headers: globalHeaders,
    prices: globalMenuPrices
  };
}

export default function DigitizerApp() {
  const [files, setFiles] = useState<File[]>([]);
  const [rows, setRows] = useState<DisplayRow[]>([]);
  const [loading, setLoading] = useState(false);
  const [selectedRows, setSelectedRows] = useState<Set<number>>(new Set());
  const [menuCountsByDate, setMenuCountsByDate] = useState<Record<string, Record<string, number>>>({});
  const [docInfoByImage, setDocInfoByImage] = useState<Record<number, DocInfo>>({});

  const [targetColumnIndices, setTargetColumnIndices] = useState<number[]>([]);
  const [columnHeaders, setColumnHeaders] = useState<string[]>([]);
  const [extractedMenuPrices, setExtractedMenuPrices] = useState<Record<string, number>>({});
  
  // ホバー用ステート
  const [hoveredRowId, setHoveredRowId] = useState<number | null>(null);
  const [hoveredCellIndex, setHoveredCellIndex] = useState<number | null>(null);
  const [previewImages, setPreviewImages] = useState<PageImage[]>([]);
  const [processingProgress, setProcessingProgress] = useState<{ current: number, total: number }>();
  const [tableZoomLevel, setTableZoomLevel] = useState(100);
  const [imageZoomLevel, setImageZoomLevel] = useState(100);
  const [zoomedImageIndex, setZoomedImageIndex] = useState<number | null>(null);
  const [debugMode, setDebugMode] = useState<boolean>(false);
  const [ocrOnlyMode, setOcrOnlyMode] = useState(false); // OCR結果のみフル表示モード

  // 一括編集モーダル用state
  const [flashColIndex, setFlashColIndex] = useState<number | null>(null);
  const [bulkModalOpen, setBulkModalOpen] = useState(false);
  const [bulkTargetColIndex, setBulkTargetColIndex] = useState<number | null>(null);
  const [bulkScope, setBulkScope] = useState<"all" | "image" | "page">("all");
  const [bulkTargetImageIndex, setBulkTargetImageIndex] = useState<number>(0);
  const [bulkTargetPage, setBulkTargetPage] = useState<number>(1);

  const getDateKey = (info?: DocInfo | null, fallbackId?: string | number) => {
    if (info?.year && info?.month && info?.day) {
      return `${info.year}-${info.month}-${info.day}`;
    }
    if (fallbackId !== undefined) {
      return `不明-${fallbackId}`;
    }
    return "不明";
  };

  const formatDateLabel = (info?: DocInfo | null, fallbackKey?: string) => {
    if (info?.year && info?.month && info?.day) {
      return `${info.year}年${info.month}月${info.day}日 (${info.dayOfWeek})`;
    }
    return fallbackKey ? `日付不明 (${fallbackKey})` : "日付不明";
  };

  const buildDocInfoByDate = (infoByImage: Record<number, DocInfo>) => {
    const byDate: Record<string, DocInfo> = {};
    Object.entries(infoByImage).forEach(([index, info]) => {
      const key = getDateKey(info, index);
      if (!byDate[key]) {
        byDate[key] = info;
        console.log(`  📅 日付マッピング: ${key} → 画像${index}:`, info);
      }
    });
    console.log(`  ✅ buildDocInfoByDate結果: ${Object.keys(byDate).length}件の日付`);
    return byDate;
  };

  const normalizeFacilityName = (facilityName: string): string => {
    if (!facilityName) return facilityName;
    // フロア情報を除去（2F, 3F, 1階, 2階, ２F, ３階 など）
    return facilityName
      .replace(/[　\s]*[0-9０-９]+[FfＦｆ階]/g, '')  // 数字+F または 数字+階
      .replace(/[　\s]*[0-9０-９]+[stndrdth]*[　\s]*floor/gi, '')  // 英語のfloor表記
      .trim();
  };

  const fetchDocInfo = async (text: string) => {
    if (!text) return null;
    try {
      const geminiRes = await fetch('/api/gemini', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ text }),
      });
      if (geminiRes.ok) {
        return await geminiRes.json();
      }
      console.warn('⚠️ Gemini API呼び出し失敗:', await geminiRes.text());
    } catch (geminiError) {
      console.warn('⚠️ Gemini APIエラー:', geminiError);
    }
    return null;
  };

  const tableScrollRef = useRef<HTMLDivElement>(null);
  const stickyHeaderRef = useRef<HTMLDivElement>(null);

  // メモリクリーンアップ: コンポーネントアンマウント時のみobject URLを解放
  useEffect(() => {
    return () => {
      previewImages.forEach(page => URL.revokeObjectURL(page.imageUrl));
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []); // 空の依存配列: アンマウント時のみクリーンアップ

  // テーブルが更新されたときにスクロール位置を左端にリセット
  useEffect(() => {
    if (rows.length > 0 && tableScrollRef.current) {
      const resetScroll = () => {
        if (tableScrollRef.current) {
          tableScrollRef.current.scrollLeft = 0;
        }
      };
      resetScroll();
      setTimeout(resetScroll, 50);
      setTimeout(resetScroll, 100);
    }
  }, [rows]);

  // --- メニュー列の動的追加 ---
  const addNewMenuColumn = () => {
    const menuName = prompt("追加するメニュー名を入力してください");
    if (!menuName) return;

    setColumnHeaders((prev) => {
      const next = [...prev];
      next.splice(next.length - 1, 0, menuName);
      return next;
    });

    setRows((prev) =>
      prev.map((row) => {
        const isHeader = row.results.every((r) => r === null);
        const nextResults = [...row.results];
        const nextColumns = [...row.columns];
        const insertIdx = nextResults.length - 1;
        nextResults.splice(insertIdx, 0, isHeader ? null : "×");
        nextColumns.splice(insertIdx, 0, isHeader ? menuName : "");
        return { ...row, results: nextResults, columns: nextColumns };
      })
    );
  };

  // --- メニュー列の削除 ---
  const removeMenuColumn = (colIndex: number) => {
    const menuName = columnHeaders[colIndex];
    if (menuName === "氏名" || menuName === "施術実施") {
      alert("この列は削除できません");
      return;
    }
    if (!confirm(`メニュー「${menuName}」を削除しますか？`)) return;

    setColumnHeaders((prev) => prev.filter((_, i) => i !== colIndex));
    setRows((prev) =>
      prev.map((row) => ({
        ...row,
        results: row.results.filter((_, i) => i !== colIndex),
        columns: row.columns.filter((_, i) => i !== colIndex),
      }))
    );
  };

  // --- メニュー列の並び替え（左へ・右へ） ---
  const moveMenuColumn = (colIndex: number, direction: "left" | "right") => {
    if (colIndex <= 0 || colIndex >= columnHeaders.length - 1) return;
    const name = columnHeaders[colIndex];
    if (name === "氏名" || name === "施術実施") return;
    const swapIndex = direction === "left" ? colIndex - 1 : colIndex + 1;
    const swapName = columnHeaders[swapIndex];
    if (swapName === "氏名" || swapName === "施術実施") return;

    const nextHeaders = [...columnHeaders];
    [nextHeaders[colIndex], nextHeaders[swapIndex]] = [nextHeaders[swapIndex], nextHeaders[colIndex]];
    setColumnHeaders(nextHeaders);

    setRows((prev) =>
      prev.map((row) => {
        const nextResults = [...row.results];
        const nextColumns = [...row.columns];
        if (nextResults.length > Math.max(colIndex, swapIndex) && nextColumns.length > Math.max(colIndex, swapIndex)) {
          [nextResults[colIndex], nextResults[swapIndex]] = [nextResults[swapIndex], nextResults[colIndex]];
          [nextColumns[colIndex], nextColumns[swapIndex]] = [nextColumns[swapIndex], nextColumns[colIndex]];
        }
        return { ...row, results: nextResults, columns: nextColumns };
      })
    );

    setFlashColIndex(swapIndex);
    setTimeout(() => setFlashColIndex(null), 700);
  };

  // --- 人の追加（指定グループに追加） ---
  const addNewPersonRowToGroup = (groupId: number) => {
    const personName = prompt("追加する氏名を入力してください");
    if (!personName) return;

    setRows((prev) => {
      const maxIdx = prev.length > 0 ? Math.max(...prev.map(r => r.rowIndex)) : 0;
      const maxRowId = prev.length > 0 ? Math.max(...prev.map(r => r.rowId ?? r.rowIndex)) : -1;
      const groupRows = prev.filter(r => (r.groupId ?? 0) === groupId);
      const firstInGroup = groupRows[0];
      const sourceImageIndex = firstInGroup?.sourceImageIndex ?? groupId;
      const sourceImageName = firstInGroup?.sourceImageName ?? previewImages[groupId]?.fileName ?? `グループ ${groupId + 1}`;
      const newRow: DisplayRow = {
        rowIndex: maxIdx + 1,
        rowId: maxRowId + 1,
        columns: columnHeaders.map((h, i) => i === 0 ? personName : ""),
        results: columnHeaders.map((h, i) => i === 0 ? null : "×"),
        sourceImageIndex,
        sourceImageName,
        groupId
      };
      return [...prev, newRow];
    });
  };

  // --- 人の削除 ---
  const removePersonRow = (rowIndex: number) => {
    const row = rows.find(r => r.rowIndex === rowIndex);
    if (!row) return;
    if (!confirm(`「${row.columns[0]}」さんの行を削除しますか？`)) return;
    setRows((prev) => prev.filter(r => r.rowIndex !== rowIndex));
  };

  // --- 画像回転 ---
  const rotateImage = (index: number) => {
    setPreviewImages((prev) => {
      const newImages = [...prev];
      newImages[index] = {
        ...newImages[index],
        rotation: (newImages[index].rotation + 90) % 360
      };
      return newImages;
    });
  };

  // --- 回転を反映した画像を生成 ---
  const createRotatedImage = async (pageImage: PageImage): Promise<Blob> => {
    if (pageImage.rotation === 0) {
      return pageImage.blob;
    }

    // 画像を読み込む
    const img = await new Promise<HTMLImageElement>((resolve, reject) => {
      const image = new Image();
      image.onload = () => resolve(image);
      image.onerror = reject;
      image.src = pageImage.imageUrl;
    });

    // Canvasを作成
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    if (!ctx) throw new Error('Canvas context not available');

    // 回転角度に応じてcanvasのサイズを設定
    const rotation = pageImage.rotation;
    if (rotation === 90 || rotation === 270) {
      canvas.width = img.height;
      canvas.height = img.width;
    } else {
      canvas.width = img.width;
      canvas.height = img.height;
    }

    // 回転を適用
    ctx.translate(canvas.width / 2, canvas.height / 2);
    ctx.rotate((rotation * Math.PI) / 180);
    ctx.drawImage(img, -img.width / 2, -img.height / 2);

    // Blobに変換
    return new Promise((resolve, reject) => {
      canvas.toBlob(
        (blob) => blob ? resolve(blob) : reject(new Error('Blob creation failed')),
        'image/png',
        0.95
      );
    });
  };

  const removePreviewImage = (index: number) => {
    setPreviewImages((prev) => {
      const target = prev[index];
      if (target) URL.revokeObjectURL(target.imageUrl);
      return prev.filter((_, i) => i !== index);
    });

    setZoomedImageIndex((prev) => {
      if (prev === null) return prev;
      if (prev === index) return null;
      if (prev > index) return prev - 1;
      return prev;
    });
  };

  // --- 解析実行 ---
  const onSubmit = async () => {
    if (previewImages.length === 0) {
      alert('画像を選択してください。');
      return;
    }

    setLoading(true);
    setRows([]);
    setSelectedRows(new Set());
    setMenuCountsByDate({});
    setDocInfoByImage({});
    setColumnHeaders([]);
    setTargetColumnIndices([]);

    try {
      let allResults: any[] = [];
      const docInfoMap: Record<number, DocInfo> = {};
      const totalPages = previewImages.length;

      setProcessingProgress({ current: 0, total: totalPages });
      let globalImageIndex = 0;
      let globalGroupId = 0;

      console.log(`\n📦 ${previewImages.length}枚の画像を処理開始`);

      // 各プレビュー画像を処理
      for (let i = 0; i < previewImages.length; i++) {
        const pageImage = previewImages[i];
        console.log(`\n📄 画像 ${i + 1}/${previewImages.length}: ${pageImage.fileName} (回転: ${pageImage.rotation}°)`);

        // 回転を反映した画像を生成
        const rotatedBlob = await createRotatedImage(pageImage);

        const formData = new FormData();
        formData.append('file', rotatedBlob, `${pageImage.fileName}.png`);

        console.log(`  🌐 API呼び出し中...`);

        // リトライ機能付きAPI呼び出し
        let res;
        let data;
        let retryCount = 0;
        const maxRetries = 5;

        while (retryCount <= maxRetries) {
          try {
            res = await fetch('/api/analyze', { method: 'POST', body: formData });

            if (res.status === 403 || res.status === 429) {
              if (retryCount < maxRetries) {
                const waitTime = (5 + retryCount * 2) * 1000;
                const errorType = res.status === 403 ? "アクセス権限/クォータエラー (403)" : "レート制限 (429)";
                console.warn(`  ⚠️ ${errorType}。${waitTime / 1000}秒待機後にリトライ... (${retryCount + 1}/${maxRetries})`);

                if (res.status === 403) {
                  await new Promise(r => setTimeout(r, waitTime + 2000));
                } else {
                  await new Promise(r => setTimeout(r, waitTime));
                }
                retryCount++;
                continue;
              } else {
                console.error(`  ❌ 最大リトライ回数に達しました`);
                if (res.status === 403) {
                  throw new Error(`APIアクセスが拒否されました (403)。\nAzure Document IntelligenceのFree Tier (F0) のクォータ制限(月間制限または同時アクセス制限)を超過した可能性があります。\n時間を空けて試すか、リソースの価格レベルを確認してください。`);
                }
                throw new Error(`API呼び出しが失敗しました: ${res.status} ${res.statusText}`);
              }
            }

            if (!res.ok) {
              throw new Error(`API呼び出しが失敗しました: ${res.status} ${res.statusText}`);
            }

            data = await res.json();
            break;
          } catch (error) {
            if (retryCount >= maxRetries) {
              throw error;
            }
            retryCount++;
            await new Promise(r => setTimeout(r, 2000));
          }
        }

        const pageMarkdown = data.analyzeResult?.content || '';
        if (pageMarkdown) {
          console.log(`  📝 Markdown抽出完了: ${pageMarkdown.length}文字`);
          const info = await fetchDocInfo(pageMarkdown);
          if (info) {
            // 施設名からフロア情報を除去
            if (info.facilityName) {
              info.facilityName = normalizeFacilityName(info.facilityName);
            }
            docInfoMap[i] = info;
            console.log('✅ 文書情報取得完了:', info);
          }
        }

        const tables: Table[] = data?.analyzeResult?.tables ?? [];
        console.log(`  📊 検出されたテーブル数: ${tables.length}`);

        // ヘッダー行で "selected"/"unselected" が出ないよう、ヘッダー判定側で除外する。
        // セル内容は正規化しない（〇/×判定で "selected"→〇 "unselected"→× として利用するため）
        tables.forEach((_tbl) => {
          /* セル内容の一括置換は行わない。buildDisplayRows でヘッダー除外と〇/×解釈を行う */
        });

        if (debugMode && tables.length > 0) {
          tables.forEach((t, idx) => {
            console.log(`    テーブル[${idx}]: ${t.rowCount}行 x ${t.columnCount}列 (合計=${t.rowCount + t.columnCount})`);
          });
        }

        const validTables = tables.filter(t => (t.rowCount + t.columnCount) > 10);
        console.log(`  ✓ 有効なテーブル数 (row+col > 10): ${validTables.length}`);

        for (const table of validTables) {
          // セルソート
          table.cells.sort((a, b) => {
            const a_y = Math.min(a.boundingRegions?.[0]?.polygon[1] ?? 0,
              a.boundingRegions?.[0]?.polygon[3] ?? 0);
            const b_y = Math.min(b.boundingRegions?.[0]?.polygon[1] ?? 0,
              b.boundingRegions?.[0]?.polygon[3] ?? 0);
            if (a_y !== b_y) return a_y - b_y;
            const a_x = Math.min(a.boundingRegions?.[0]?.polygon[0] ?? 0,
              a.boundingRegions?.[0]?.polygon[6] ?? 0);
            const b_x = Math.min(b.boundingRegions?.[0]?.polygon[0] ?? 0,
              b.boundingRegions?.[0]?.polygon[6] ?? 0);
            return a_x - b_x;
          });

          console.log(`  🔧 buildDisplayRows実行中... (${table.rowCount}行 x ${table.columnCount}列)`);
          try {
            const buildResult = await buildDisplayRows(table, pageImage.imageUrl, globalImageIndex, pageImage.fileName, globalGroupId, rotatedBlob, debugMode);

            if (buildResult.displayRows.length > 0 && buildResult.headers.length > 0) {
              console.log(`  ✅ buildDisplayRows完了: ${buildResult.displayRows.length}行、${buildResult.headers.length}列のデータ`);
              allResults.push(buildResult);
              globalGroupId++;
            } else if (allResults.length > 0) {
              // ヘッダー行なしページ: 直前の成功結果の列構成でリトライ
              console.warn(`  ⚠️ buildDisplayRows結果が空 → フォールバックでリトライ`);
              const ref = allResults[allResults.length - 1];
              const fallbackResult = await buildDisplayRows(table, pageImage.imageUrl, globalImageIndex, pageImage.fileName, globalGroupId, rotatedBlob, debugMode, { indices: ref.indices, headers: ref.headers });
              if (fallbackResult.displayRows.length > 0 && fallbackResult.headers.length > 0) {
                console.log(`  ✅ フォールバック成功: ${fallbackResult.displayRows.length}行`);
                allResults.push(fallbackResult);
                globalGroupId++;
              } else {
                console.warn(`  ⚠️ フォールバックも失敗 → スキップ`);
              }
            } else {
              console.warn(`  ⚠️ buildDisplayRows結果が空: displayRows=${buildResult.displayRows.length}, headers=${buildResult.headers.length}`);
              console.warn(`  → このテーブルはスキップされます`);
            }
          } catch (error) {
            console.error(`  ❌ buildDisplayRowsエラー:`, error);
            console.error(`  → このテーブルはスキップして続行します`);
          }
        }

        globalImageIndex++;
        setProcessingProgress({ current: i + 1, total: totalPages });

        // レート制限（1000ms待機）
        if (i < previewImages.length - 1) {
          await new Promise(r => setTimeout(r, 1000));
        }
      }

      console.log(`\n📦 全画像処理完了。allResults配列: ${allResults.length}個のテーブル結果`);

      if (allResults.length === 0) {
        console.error('❌ 有効なテーブルが1つも検出されませんでした');
        alert('有効なテーブルが検出されませんでした。\n画像に「氏名」列を含むテーブルが存在することを確認してください。');
        return;
      }

      // 施設名がないPDFは同じ書類の続きとして前のPDFと同じ情報（日付含む）を適用
      let lastValidInfo: DocInfo | null = null;
      for (let idx = 0; idx < previewImages.length; idx++) {
        const current = docInfoMap[idx];
        const hasFacilityName = !!current?.facilityName;
        if (hasFacilityName) {
          lastValidInfo = current;
        } else if (lastValidInfo) {
          docInfoMap[idx] = { ...lastValidInfo };
          console.log(`  📋 画像${idx}: 施設名なし → 前のPDFと同じ情報（日付含む）を適用`);
        }
      }

      setDocInfoByImage(docInfoMap);

      // ===== 共通マージ処理 =====
      console.log(`\n🔗 マージ処理開始: ${allResults.length}個の結果を結合`);
      const merged = mergeDisplayResults(allResults);

      if (merged) {
        console.log(`✅ マージ完了: 最終的に${merged.displayRows.length}行、${merged.headers.length}列のデータ`);
        console.log(`   ヘッダー: ${merged.headers.join(', ')}`);
        if (merged.prices && Object.keys(merged.prices).length > 0) {
          console.log(`   💰 抽出された価格: ${Object.entries(merged.prices).map(([m, p]) => `${m}(¥${p})`).join(', ')}`);
        }

        const finalRows = merged.displayRows.map((row: DisplayRow, idx: number) => ({
          ...row,
          rowIndex: idx,
          rowId: idx
        }));

        setRows(finalRows);
        setTargetColumnIndices(merged.indices);
        setColumnHeaders(merged.headers);
        setExtractedMenuPrices(merged.prices || {});
      } else {
        console.warn('⚠️ マージ結果がnullです');
        alert('データのマージに失敗しました。');
      }

    } catch (error) {
      console.error('❌ 解析エラー:', error);

      let errorMessage = '解析中にエラーが発生しました。';
      if (error instanceof Error) {
        errorMessage += `\n\nエラー詳細: ${error.message}`;

        // レート制限エラーの場合
        if (error.message.includes('403') || error.message.includes('429')) {
          errorMessage += '\n\n⚠️ Azure APIのレート制限に達した可能性があります。';
          errorMessage += '\n数分待ってから再度お試しください。';
          errorMessage += '\nまたは、一度に処理するファイル数を減らしてください。';
        }
      }

      alert(errorMessage);
    } finally {
      setLoading(false);
      setProcessingProgress(undefined);
    }
  };

  const toggleResult = (rowId: number, colIndex: number) => {
    setRows((prev) => {
      const index = prev.findIndex((r) => (r.rowId ?? r.rowIndex) === rowId);
      if (index === -1) return prev;
      const row = prev[index];
      const nextResults = [...row.results];
      const current = nextResults[colIndex];
      nextResults[colIndex] = current === "〇" ? "×" : "〇";
      const next = [...prev];
      next[index] = { ...row, results: nextResults };
      return next;
    });
  };

  // 一括更新ロジック
  const executeBulkUpdate = (action: "ok" | "ng" | "toggle") => {
    if (bulkTargetColIndex === null) return;

    setRows(prev => prev.map(row => {
      // ヘッダー行はスキップ
      if (row.results.every(r => r === null)) return row;

      // 対象判定
      let isTarget = false;
      if (bulkScope === "all") {
        isTarget = true;
      } else if (bulkScope === "image") {
        if (row.sourceImageIndex === bulkTargetImageIndex) isTarget = true;
      } else if (bulkScope === "page") {
        // sourceImageNameからページ番号を推測するか、
        // PageImageのpageNumberを使うにはrowに紐づけが必要。
        // 簡易的に previewImages[row.sourceImageIndex].pageNumber を参照
        if (row.sourceImageIndex !== undefined) {
          const imgInfo = previewImages[row.sourceImageIndex];
          if (imgInfo && imgInfo.pageNumber === bulkTargetPage) isTarget = true;
        }
      }

      if (!isTarget) return row;

      const nextResults = [...row.results];
      const currentVal = nextResults[bulkTargetColIndex];

      let nextVal = currentVal;
      if (action === "ok") nextVal = "〇";
      if (action === "ng") nextVal = "×";
      if (action === "toggle") {
        nextVal = (currentVal === "〇") ? "×" : "〇";
      }

      nextResults[bulkTargetColIndex] = nextVal;
      return { ...row, results: nextResults };
    }));
    setBulkModalOpen(false);
  };

  const toggleRowSelection = (rowIndex: number) => {
    setSelectedRows((prev) => {
      const newSet = new Set(prev);
      if (newSet.has(rowIndex)) newSet.delete(rowIndex);
      else newSet.add(rowIndex);
      return newSet;
    });
  };

  const countMenuResults = () => {
    console.log("🔍 集計開始 - columnHeaders:", columnHeaders);

    const menuResultIndices = columnHeaders
      .map((name, index) => ({ name, index }))
      .filter((x) => x.name !== "氏名" && x.name !== "施術実施" && !x.name.startsWith("※"));

    console.log("🔍 メニュー項目:", menuResultIndices.map(x => x.name));

    const createEmptyCounts = () => {
      const counts: Record<string, number> = {};
      menuResultIndices.forEach(({ name }) => { counts[name] = 0; });
      return counts;
    };

    const countsByDate: Record<string, Record<string, number>> = {};

    rows.forEach((row) => {
      if (row.results.every((res) => res === null)) return;
      const shijitsuResult = row.results[row.results.length - 1];
      if (shijitsuResult !== "〇") return;

      const info = row.sourceImageIndex !== undefined ? docInfoByImage[row.sourceImageIndex] : undefined;
      const dateKey = getDateKey(info, row.sourceImageIndex ?? "unknown");
      if (!countsByDate[dateKey]) countsByDate[dateKey] = createEmptyCounts();

      menuResultIndices.forEach(({ name, index }) => {
        if (row.results[index] === "〇") countsByDate[dateKey][name]++;
      });
    });

    console.log("🔍 集計結果 - countsByDate:", countsByDate);
    setMenuCountsByDate(countsByDate);
  };

  const onExportExcel = async () => {
    setLoading(true);
    try {
      console.log("🔍 Excel Export - docInfoByImage:", docInfoByImage);

      const countsByDate = (() => {
        const menuResultIndices = columnHeaders
          .map((name, index) => ({ name, index }))
          .filter((x) => x.name !== "氏名" && x.name !== "施術実施" && !x.name.startsWith("※"));

        const createEmptyCounts = () => {
          const counts: Record<string, number> = {};
          menuResultIndices.forEach(({ name }) => { counts[name] = 0; });
          return counts;
        };

        const counts: Record<string, Record<string, number>> = {};
        rows.forEach((row) => {
          if (row.results.every((res) => res === null)) return;
          const shijitsuResult = row.results[row.results.length - 1];
          if (shijitsuResult !== "〇") return;
          const info = row.sourceImageIndex !== undefined ? docInfoByImage[row.sourceImageIndex] : undefined;
          const dateKey = getDateKey(info, row.sourceImageIndex ?? "unknown");

          console.log(`  行 ${row.rowIndex}: sourceImageIndex=${row.sourceImageIndex}, info=`, info, `, dateKey=${dateKey}`);

          if (!counts[dateKey]) counts[dateKey] = createEmptyCounts();
          menuResultIndices.forEach(({ name, index }) => {
            if (row.results[index] === "〇") counts[dateKey][name]++;
          });
        });

        console.log("🔍 集計結果 countsByDate:", counts);
        return counts;
      })();

      let dateKeys = Object.keys(countsByDate);
      if (dateKeys.length === 0) {
        alert("集計データがありません。先に集計確定を行ってください。");
        return;
      }

      const docInfoByDate = buildDocInfoByDate(docInfoByImage);

      // 日付キーを年月日順にソート
      dateKeys = dateKeys.sort((a, b) => {
        const infoA = docInfoByDate[a];
        const infoB = docInfoByDate[b];

        // 両方とも有効な日付情報がある場合
        if (infoA?.year && infoA?.month && infoA?.day && infoB?.year && infoB?.month && infoB?.day) {
          const dateA = new Date(parseInt(infoA.year), parseInt(infoA.month) - 1, parseInt(infoA.day));
          const dateB = new Date(parseInt(infoB.year), parseInt(infoB.month) - 1, parseInt(infoB.day));
          return dateA.getTime() - dateB.getTime();
        }

        // 日付情報がない場合は文字列比較
        return a.localeCompare(b);
      });

      console.log(`📊 Excel出力データ作成開始:`);
      console.log(`   検出された日付数: ${dateKeys.length}`);
      console.log(`   ソート前の日付キー:`, Object.keys(countsByDate));
      console.log(`   ソート後の日付キー:`, dateKeys);
      console.log(`   docInfoByDate:`, docInfoByDate);

      // すべての日付データを配列にまとめる
      const dateDataList = dateKeys.map((dateKey, idx) => {
        const rawCounts = countsByDate[dateKey];
        const sentCounts: Record<string, number> = {};
        const sentUnitPrices: Record<string, number> = {};

        console.log(`\n📅 [${idx + 1}/${dateKeys.length}] 日付キー: "${dateKey}"`);
        console.log(`   rawCounts:`, rawCounts);

        Object.entries(rawCounts).forEach(([rawName, count]) => {
          // メニュー名から価格を抽出（例: "カット ¥1,800" → カット, 1800）
          const match = rawName.match(/^(.*?)([\s\u3000]*[¥￥]\s*([\d,]+))?$/);
          if (match) {
            let name = match[1].trim();
            if (name === "顔剃り") name = "顔そり";
            const priceStr = match[3] ? match[3].replace(/,/g, '') : "";
            let price = priceStr ? parseInt(priceStr, 10) : 0;

            // 白髪染めをカラー（白髪染め）に統合
            if (name === "白髪染め") {
              name = "カラー（白髪染め）";
            }

            const finalName = name || rawName;

            // 価格の優先順位: 1. 抽出された価格 2. 埋め込み価格 3. デフォルト価格
            if (price === 0) {
              // 白髪染めの場合、カラーの価格を使用
              if (finalName === "カラー（白髪染め）" && extractedMenuPrices["カラー"]) {
                price = extractedMenuPrices["カラー"];
                console.log(`   💎 メニュー "${finalName}": カラーの価格を使用 ¥${price}`);
              } else if (extractedMenuPrices[finalName]) {
                price = extractedMenuPrices[finalName];
                console.log(`   💎 メニュー "${finalName}": 表から抽出した価格を使用 ¥${price}`);
              } else if (DEFAULT_MENU_PRICES[finalName]) {
                price = DEFAULT_MENU_PRICES[finalName];
                console.log(`   💰 メニュー "${finalName}": デフォルト価格を使用 ¥${price}`);
              } else if (finalName === "カラー（白髪染め）" && DEFAULT_MENU_PRICES["カラー"]) {
                price = DEFAULT_MENU_PRICES["カラー"];
                console.log(`   💰 メニュー "${finalName}": カラーのデフォルト価格を使用 ¥${price}`);
              }
            } else {
              console.log(`   📝 メニュー "${finalName}": 埋め込み価格を使用 ¥${price}`);
            }

            // カウントを加算（既存の値がある場合）
            sentCounts[finalName] = (sentCounts[finalName] || 0) + count;
            // 価格は最初に設定された値を保持
            if (!sentUnitPrices[finalName]) {
              sentUnitPrices[finalName] = price;
            }

            if (price === 0) {
              console.log(`   ⚠️ メニュー "${rawName}" → 価格が0円（すべての価格ソースで未定義）`);
            } else {
              console.log(`   ✓ メニュー "${rawName}" → "${finalName}": ${count}人 x ¥${price} = ¥${count * price}`);
            }
          } else {
            // マッチ失敗した場合も抽出価格とデフォルト価格を試す
            const price = extractedMenuPrices[rawName] || DEFAULT_MENU_PRICES[rawName] || 0;
            sentCounts[rawName] = count;
            sentUnitPrices[rawName] = price;

            if (price > 0) {
              const source = extractedMenuPrices[rawName] ? "表から抽出" : "デフォルト";
              console.log(`   💰 メニュー "${rawName}": ${source}価格を使用 ¥${price}`);
            } else {
              console.log(`   ❌ メニュー "${rawName}": 価格情報なし → ¥0`);
            }
          }
        });

        const info = docInfoByDate[dateKey];
        let reiwaYearStr = "";
        if (info?.year) {
          const yearNum = parseInt(info.year);
          if (!isNaN(yearNum)) reiwaYearStr = `令和${yearNum - 2018}年`;
        }

        const dateData = {
          counts: sentCounts,
          unitPrices: sentUnitPrices,
          facility: info?.facilityName || "",
          reiwaYear: reiwaYearStr,
          month: info?.month || "",
          day: info?.day || "",
          weekday: info?.dayOfWeek || ""
        };

        console.log(`   📋 送信データ:`, dateData);
        console.log(`   📊 sentCounts件数: ${Object.keys(sentCounts).length}, sentUnitPrices件数: ${Object.keys(sentUnitPrices).length}`);

        // データが空の場合は警告
        if (Object.keys(sentCounts).length === 0) {
          console.warn(`   ⚠️ 警告: 日付 "${dateKey}" のsentCountsが空です！`);
        }

        return dateData;
      });

      console.log(`\n✅ 最終送信データ (${dateDataList.length}件):`, dateDataList);

      console.log("📋 Export Payload Preview:", { dateDataList });

      // 申し込みカラムのメニューのみ請求書に含める（氏名・施術実施・※注釈を除く。価格表記を除いたベース名で渡す）
      const allowedMenus = columnHeaders
        .filter((h) => h !== "氏名" && h !== "施術実施" && !h.startsWith("※"))
        .map((h) => {
          const m = h.match(/^(.*?)([\s\u3000]*[¥￥].*)?$/);
          return m ? m[1].trim() : h;
        });

      const res = await fetch("/api/export", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ dateDataList, allowedMenus }),
      });

      if (!res.ok) {
        const errorText = await res.text();
        console.error("Excel export failed:", errorText);
        throw new Error(`Export failed: ${errorText}`);
      }

      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      const firstInfo = docInfoByDate[dateKeys[0]];
      const fileDate = firstInfo?.year && firstInfo?.month && firstInfo?.day
        ? `${firstInfo.year}${firstInfo.month}${firstInfo.day}`
        : dateKeys[0].replace(/[^\d]/g, '') || "unknown";
      a.href = url;
      a.download = `${firstInfo?.facilityName || 'order'}_${fileDate}.xlsx`;
      a.click();
    } catch (err) {
      alert("処理中にエラーが発生しました");
    } finally {
      setLoading(false);
    }
  };

  const docInfoByDate = buildDocInfoByDate(docInfoByImage);

  return (
    <main style={{ width: "100%", maxWidth: "100vw", margin: "0", padding: "24px", fontFamily: "sans-serif", backgroundColor: "#f9fafb", color: "#111827", minHeight: "100vh", overflowX: "auto", boxSizing: "border-box" }}>
      <div style={{ backgroundColor: "white", color: "#111827", padding: "24px", borderRadius: "12px", boxShadow: "0 1px 3px rgba(0,0,0,0.1)", marginBottom: "24px" }}>
        <h1 style={{ fontSize: "24px", fontWeight: "bold", marginBottom: "20px", color: "#111827" }}>申込書清書システム</h1>

        <div style={{ display: "flex", flexWrap: "wrap", gap: "12px", alignItems: "center", marginBottom: "24px" }}>

          <label style={{ display: "flex", alignItems: "center", gap: "6px", fontSize: "14px", fontWeight: "600", color: "#111827", backgroundColor: "#f3f4f6", padding: "8px 12px", borderRadius: "6px", cursor: "pointer", border: "1px solid #d1d5db" }}>
            <input
              type="checkbox"
              checked={debugMode}
              onChange={(e) => setDebugMode(e.target.checked)}
            />
            🛠️ デバッグモード
          </label>

          <input
            type="file"
            accept="image/*,application/pdf"
            multiple
            style={{ padding: "8px", border: "1px solid #d1d5db", borderRadius: "6px", flex: "1", minWidth: "200px" }}
            onChange={async (e) => {
              const fileList = e.target.files;
              if (fileList && fileList.length > 0) {
                const filesArray = Array.from(fileList);
                // 既存のファイルに新しいファイルを追加
                setFiles(prev => [...prev, ...filesArray]);
                setLoading(true);

                try {
                  // 既存の画像は保持し、新しい画像のみを処理
                  const newImages: PageImage[] = [];

                  for (const file of filesArray) {
                    if (file.type === 'application/pdf') {
                      // PDFを画像に変換
                      const pageImgs = await convertPdfToImages(file);
                      // ファイル名を更新
                      pageImgs.forEach(img => {
                        img.fileName = `${file.name} - ページ ${img.pageNumber}`;
                      });
                      newImages.push(...pageImgs);
                    } else {
                      // 画像ファイル
                      const imageUrl = URL.createObjectURL(file);
                      const img = await new Promise<HTMLImageElement>((resolve, reject) => {
                        const image = new Image();
                        image.onload = () => resolve(image);
                        image.onerror = reject;
                        image.src = imageUrl;
                      });

                      newImages.push({
                        blob: file,
                        imageUrl: imageUrl,
                        pageNumber: 1,
                        width: img.width,
                        height: img.height,
                        rotation: 0,
                        fileName: file.name
                      });
                    }
                  }

                  // 既存の画像に新しい画像を追加
                  setPreviewImages(prev => [...prev, ...newImages]);
                } catch (error) {
                  console.error('プレビュー生成エラー:', error);
                  alert('プレビューの生成中にエラーが発生しました。');
                } finally {
                  setLoading(false);
                }
              }
              // 同じファイルを再度選択できるように、inputの値をリセット
              e.target.value = '';
            }}
          />
          <button onClick={onSubmit} disabled={loading} style={{ padding: "10px 20px", backgroundColor: loading ? "#9ca3af" : "#2563eb", color: "white", border: "none", borderRadius: "6px", cursor: "pointer", fontWeight: "600" }}>
            {loading ? "解析中..." : "アップロード & 解析"}
          </button>
          <button onClick={addNewMenuColumn} disabled={rows.length === 0} style={{ padding: "10px 20px", backgroundColor: "#8b5cf6", color: "white", border: "none", borderRadius: "6px", cursor: "pointer", fontWeight: "600" }}>メニュー追加</button>
          <button onClick={countMenuResults} style={{ padding: "10px 20px", backgroundColor: "#059669", color: "white", border: "none", borderRadius: "6px", cursor: "pointer", fontWeight: "600" }}>データ確定</button>
        </div>

        {processingProgress && (
          <div style={{
            padding: '8px 16px',
            backgroundColor: '#eff6ff',
            borderRadius: '6px',
            fontSize: '14px',
            color: '#1e40af',
            marginBottom: '16px'
          }}>
            処理中: {processingProgress.current} / {processingProgress.total} ページ
          </div>
        )}

        {Object.keys(docInfoByDate).length > 0 && (
          <div style={{ marginBottom: "24px", padding: "16px", backgroundColor: "#f0fdf4", borderRadius: "8px", border: "1px solid #bbf7d0", display: "flex", flexDirection: "column", gap: "12px", color: "#166534" }}>
            {Object.entries(docInfoByDate).map(([dateKey, info]) => (
              <div key={dateKey} style={{ display: "flex", gap: "24px", flexWrap: "wrap" }}>
                <div><strong>施設名:</strong> {info.facilityName || "(未取得)"}</div>
                <div><strong>施術日:</strong> {formatDateLabel(info, dateKey)}</div>
              </div>
            ))}
          </div>
        )}


      </div>

      <div
        style={{
          display: "flex",
          gap: "24px",
          overflowX: "auto",
          overflowY: "visible",
          scrollSnapType: "none",
          scrollBehavior: "smooth",
          marginBottom: "24px",
          WebkitOverflowScrolling: "touch",
          paddingBottom: "8px",
          width: "100%",
          minWidth: 0,
          alignItems: "flex-start",
          flexDirection: ocrOnlyMode ? "column" : "row"
        }}
      >
        {/* 左側: 画像プレビュー（入れ替え後） */}
        {!ocrOnlyMode && (
          <div
            style={{
              flex: "0 0 auto",
              minWidth: "600px",
              maxWidth: "calc(50% - 12px)",
              backgroundColor: "white",
              color: "#111827",
              padding: "24px",
              borderRadius: "12px",
              boxShadow: "0 1px 3px rgba(0,0,0,0.1)",
              display: "flex",
              flexDirection: "column",
              flexShrink: 0
            }}
          >
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "16px", flexShrink: 0 }}>
              <h2 style={{ fontSize: "18px", fontWeight: "bold", margin: 0, color: "#111827" }}>
                画像プレビュー {previewImages.length > 0 ? `(${previewImages.length}枚)` : ""}
              </h2>
              <div style={{ display: "flex", gap: "8px", alignItems: "center" }}>
                <button
                  onClick={() => setImageZoomLevel(prev => Math.max(30, prev - 10))}
                  style={{ padding: "4px 12px", backgroundColor: "#f3f4f6", color: "#374151", border: "1px solid #d1d5db", borderRadius: "4px", fontSize: "14px", cursor: "pointer", fontWeight: "bold" }}
                  title="縮小"
                >−</button>
                <span style={{ padding: "4px 12px", backgroundColor: "#f9fafb", border: "1px solid #e5e7eb", borderRadius: "4px", fontSize: "12px", minWidth: "50px", textAlign: "center", display: "inline-block", color: "#111827" }}>
                  {imageZoomLevel}%
                </span>
                <button
                  onClick={() => setImageZoomLevel(prev => Math.min(300, prev + 10))}
                  style={{ padding: "4px 12px", backgroundColor: "#f3f4f6", color: "#374151", border: "1px solid #d1d5db", borderRadius: "4px", fontSize: "14px", cursor: "pointer", fontWeight: "bold" }}
                  title="拡大"
                >＋</button>
              </div>
            </div>
            <div
              style={{
                flex: "1",
                overflowX: "hidden",
                overflowY: "auto",
                minHeight: 0,
                width: "100%",
                maxHeight: "800px",
                WebkitOverflowScrolling: "touch"
              }}
            >
              {previewImages.length > 0 ? (
                <div style={{ display: "flex", flexDirection: "column", gap: "16px", paddingBottom: "8px" }}>
                  {previewImages.map((page, index) => (
                    <div
                      key={index}
                      style={{
                        flex: "0 0 auto",
                        border: "1px solid #e5e7eb",
                        borderRadius: "8px",
                        overflow: "hidden",
                        transition: "box-shadow 0.2s",
                        position: "relative"
                      }}
                      onMouseEnter={(e) => {
                        e.currentTarget.style.boxShadow = "0 4px 8px rgba(0,0,0,0.2)";
                      }}
                      onMouseLeave={(e) => {
                        e.currentTarget.style.boxShadow = "none";
                      }}
                    >
                      <button
                        onClick={(e) => {
                          e.stopPropagation();
                          removePreviewImage(index);
                        }}
                        style={{
                          position: "absolute",
                          top: "8px",
                          left: "8px",
                          backgroundColor: "#ffffff",
                          color: "#111827",
                          border: "1px solid #d1d5db",
                          borderRadius: "4px",
                          padding: "6px 10px",
                          fontSize: "12px",
                          cursor: "pointer",
                          fontWeight: "600",
                          boxShadow: "0 2px 4px rgba(0,0,0,0.1)",
                          zIndex: 10
                        }}
                        title="画像を削除"
                      >
                        🗑️ 削除
                      </button>
                      <button
                        onClick={(e) => {
                          e.stopPropagation();
                          rotateImage(index);
                        }}
                        style={{
                          position: "absolute",
                          top: "8px",
                          right: "8px",
                          backgroundColor: "#ffffff",
                          color: "#111827",
                          border: "1px solid #d1d5db",
                          borderRadius: "4px",
                          padding: "6px 12px",
                          fontSize: "12px",
                          cursor: "pointer",
                          fontWeight: "600",
                          boxShadow: "0 2px 4px rgba(0,0,0,0.1)",
                          zIndex: 10
                        }}
                        title="90度回転"
                      >
                        🔄 回転
                      </button>

                      <div
                        onClick={() => setZoomedImageIndex(index)}
                        style={{ cursor: "zoom-in", position: "relative" }}
                      >
                        <img
                          src={page.imageUrl}
                          alt={`Image ${index + 1}`}
                          style={{
                            height: `${600 * imageZoomLevel / 100}px`,
                            width: "auto",
                            display: "block",
                            objectFit: "contain",
                            transform: `rotate(${page.rotation}deg)`,
                            transition: "transform 0.3s ease"
                          }}
                        />
                        {/* Coordinate Highlights Overlay */}
                        <svg 
                          style={{ 
                            position: "absolute", top: 0, left: 0, width: "100%", height: "100%", 
                            pointerEvents: "none", transform: `rotate(${page.rotation}deg)` 
                          }}
                          viewBox={`0 0 ${page.width || 1} ${page.height || 1}`}
                        >
                          {rows.map((row) => {
                            if (row.sourceImageIndex !== index) return null;
                            const isRowHovered = hoveredRowId === row.rowId;
                            return (
                              <g key={row.rowId}>
                                {/* 1. 行全体のハイライト (氏名セルなど) */}
                                {row.polygon && (
                                  <polygon
                                    points={row.polygon.join(',')}
                                    fill={isRowHovered ? "rgba(37, 99, 235, 0.2)" : "rgba(37, 99, 235, 0.05)"}
                                    stroke={isRowHovered ? "#2563eb" : "transparent"}
                                    strokeWidth="2"
                                  />
                                )}
                                
                                {/* 2. 個別セルのハイライト */}
                                {isRowHovered && hoveredCellIndex !== null && row.cellPolygons?.[hoveredCellIndex] && (
                                  <polygon
                                    points={row.cellPolygons[hoveredCellIndex].join(',')}
                                    fill="rgba(249, 158, 11, 0.4)"
                                    stroke="#f59e0b"
                                    strokeWidth="3"
                                  />
                                )}
                              </g>
                            );
                          })}
                        </svg>
                      </div>

                      <div style={{
                        padding: "8px",
                        backgroundColor: "#f3f4f6",
                        textAlign: "center",
                        fontSize: "11px",
                        color: "#111827"
                      }}>
                        <div style={{ fontWeight: "600", marginBottom: "4px" }}>
                          {page.fileName}
                        </div>
                        <div style={{ color: "#4b5563", fontSize: "10px" }}>
                          回転: {page.rotation}°
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              ) : (
                <div style={{
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  height: "100%",
                  color: "#4b5563",
                  fontSize: "14px"
                }}>
                  画像を選択してください
                </div>
              )}
            </div>
          </div>
        )}

        {/* 右側: OCR結果テーブル（入れ替え後） */}
        {rows.length > 0 && (
          <div
            style={{
              flex: ocrOnlyMode ? "1 1 100%" : "0 0 auto",
              minWidth: ocrOnlyMode ? "100%" : "400px",
              maxWidth: ocrOnlyMode ? "100%" : "calc(50% - 12px)",
              backgroundColor: "white",
              color: "#111827",
              padding: "24px",
              borderRadius: "12px",
              boxShadow: "0 1px 3px rgba(0,0,0,0.1)",
              display: "flex",
              flexDirection: "column",
              flexShrink: 0,
              direction: "ltr",
              width: ocrOnlyMode ? "100%" : "auto"
            }}
          >
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "16px", flexShrink: 0 }}>
              <h2 style={{ fontSize: "18px", fontWeight: "bold", margin: 0, color: "#111827" }}>OCR結果テーブル</h2>
              <div style={{ display: "flex", gap: "8px" }}>
                {/* OCRのみ表示/通常表示切り替えボタン */}
                <button
                  onClick={() => setOcrOnlyMode(prev => !prev)}
                  style={{
                    padding: "6px 12px",
                    backgroundColor: ocrOnlyMode ? "#2563eb" : "#f3f4f6",
                    color: ocrOnlyMode ? "white" : "#374151",
                    border: "1px solid #d1d5db",
                    borderRadius: "4px",
                    fontSize: "12px",
                    cursor: "pointer",
                    fontWeight: "600"
                  }}
                  title={ocrOnlyMode ? "画像を表示する" : "OCR結果のみ大きく表示する"}
                >
                  {ocrOnlyMode ? "🔲 通常表示" : "📋 OCRのみ表示"}
                </button>
                {/* 拡大縮小ボタン */}
                <button
                  onClick={() => setTableZoomLevel(prev => Math.max(50, prev - 10))}
                  style={{
                    padding: "4px 12px",
                    backgroundColor: "#f3f4f6",
                    color: "#374151",
                    border: "1px solid #d1d5db",
                    borderRadius: "4px",
                    fontSize: "14px",
                    cursor: "pointer",
                    fontWeight: "bold"
                  }}
                  title="縮小"
                >
                  −
                </button>
                <span style={{
                  padding: "4px 12px",
                  backgroundColor: "#f9fafb",
                  border: "1px solid #e5e7eb",
                  borderRadius: "4px",
                  fontSize: "12px",
                  minWidth: "50px",
                  textAlign: "center",
                  display: "inline-block",
                  color: "#111827"
                }}>
                  {tableZoomLevel}%
                </span>
                <button
                  onClick={() => setTableZoomLevel(prev => Math.min(200, prev + 10))}
                  style={{
                    padding: "4px 12px",
                    backgroundColor: "#f3f4f6",
                    color: "#374151",
                    border: "1px solid #d1d5db",
                    borderRadius: "4px",
                    fontSize: "14px",
                    cursor: "pointer",
                    fontWeight: "bold"
                  }}
                  title="拡大"
                >
                  ＋
                </button>
              </div>
            </div>
            {/* 常に表示されるスティッキーヘッダー（スクロールエリア外） */}
            {rows.length > 0 && (() => {
              const firstHeaderRow = rows.find(r => r.results.every(v => v === null));
              if (!firstHeaderRow) return null;
              return (
                <div
                  ref={stickyHeaderRef}
                  style={{
                    overflowX: "hidden",
                    border: "1px solid #e5e7eb",
                    borderBottom: "none",
                    borderRadius: "8px 8px 0 0",
                    zoom: tableZoomLevel / 100,
                    flexShrink: 0,
                  }}
                >
                  <RowView
                    row={firstHeaderRow}
                    onToggle={() => {}}
                    onHeaderDelete={removeMenuColumn}
                    onHeaderMove={moveMenuColumn}
                    onHeaderBulkEdit={(colIndex) => {
                      setBulkTargetColIndex(colIndex);
                      setBulkModalOpen(true);
                    }}
                    onRowDelete={() => {}}
                    onRowClick={() => {}}
                    isSelected={false}
                    prices={{ ...DEFAULT_MENU_PRICES, ...extractedMenuPrices }}
                    flashColIndex={flashColIndex}
                    onMouseEnterRow={() => {}}
                    onMouseLeaveRow={() => {}}
                    onMouseEnterCell={() => {}}
                    onMouseLeaveCell={() => {}}
                  />
                </div>
              );
            })()}
            <div
              ref={tableScrollRef}
              onScroll={() => {
                if (stickyHeaderRef.current && tableScrollRef.current) {
                  stickyHeaderRef.current.scrollLeft = tableScrollRef.current.scrollLeft;
                }
              }}
              style={{
                overflowX: "auto",
                overflowY: "auto",
                border: "1px solid #e5e7eb",
                borderTop: "none",
                borderRadius: "0 0 8px 8px",
                flex: "1",
                maxHeight: ocrOnlyMode ? "calc(100vh - 200px)" : "600px",
                width: "100%",
                WebkitOverflowScrolling: "touch",
                position: "relative",
                direction: "ltr",
                textAlign: "left"
              }}
            >
              <div style={{
                display: "flex",
                flexDirection: "column",
                width: "max-content",
                minWidth: "100%",
                marginLeft: 0,
                paddingLeft: 0,
                zoom: tableZoomLevel / 100,
              }}>
                {(() => {
                  // groupIdごとにグループ化
                  const groupedByGroupId: Record<number, DisplayRow[]> = {};
                  rows.forEach(row => {
                    const gId = row.groupId ?? 0;
                    if (!groupedByGroupId[gId]) {
                      groupedByGroupId[gId] = [];
                    }
                    groupedByGroupId[gId].push(row);
                  });

                  const groupIds = Object.keys(groupedByGroupId).map(Number).sort((a, b) => a - b);

                  return groupIds.map((gId) => {
                    const allGroupRows = groupedByGroupId[gId];
                    const imageName = allGroupRows[0]?.sourceImageName || `グループ ${gId + 1}`;

                    // ヘッダー行を1つだけに絞る（最初のヘッダー行のみ保持）
                    let foundFirstHeader = false;
                    const groupRows = allGroupRows.filter((row) => {
                      const isHeaderRow = row.results.every((r) => r === null);

                      if (isHeaderRow) {
                        if (foundFirstHeader) return false; // 重複ヘッダーをスキップ
                        foundFirstHeader = true;
                        return true;
                      }

                      return true; // データ行は常に表示
                    });

                    return (
                      <div
                        key={gId}
                        style={{
                          border: "2px solid #3b82f6",
                          borderRadius: "8px",
                          marginBottom: "16px",
                        }}
                      >
                        {/* 画像名ヘッダー ＋ 人を追加ボタン */}
                        <div style={{
                          backgroundColor: "#3b82f6",
                          color: "white",
                          padding: "8px 16px",
                          fontSize: "13px",
                          fontWeight: "600",
                          display: "flex",
                          justifyContent: "space-between",
                          alignItems: "center",
                          borderRadius: "6px 6px 0 0",
                        }}>
                          <span>📄 {imageName}</span>
                          <button
                            onClick={() => addNewPersonRowToGroup(gId)}
                            style={{
                              padding: "6px 12px",
                              backgroundColor: "rgba(255,255,255,0.2)",
                              color: "white",
                              border: "1px solid rgba(255,255,255,0.5)",
                              borderRadius: "6px",
                              cursor: "pointer",
                              fontWeight: "600",
                              fontSize: "12px"
                            }}
                          >
                            人を追加
                          </button>
                        </div>
                        {/* 行データ（ヘッダー行は外のstickyHeaderで表示するためスキップ） */}
                        <div>
                          {groupRows.filter(row => !row.results.every(v => v === null)).map((row) => (
                              <RowView
                                key={row.rowId ?? row.rowIndex}
                                row={row}
                                onToggle={toggleResult}
                                onHeaderDelete={removeMenuColumn}
                                onHeaderMove={moveMenuColumn}
                                onHeaderBulkEdit={(colIndex) => {
                                  setBulkTargetColIndex(colIndex);
                                  setBulkModalOpen(true);
                                }}
                                onRowDelete={removePersonRow}
                                onRowClick={toggleRowSelection}
                                isSelected={selectedRows.has(row.rowIndex)}
                                prices={{ ...DEFAULT_MENU_PRICES, ...extractedMenuPrices }}
                                flashColIndex={flashColIndex}
                                onMouseEnterRow={() => setHoveredRowId(row.rowId ?? row.rowIndex)}
                                onMouseLeaveRow={() => setHoveredRowId(null)}
                                onMouseEnterCell={(colIndex) => setHoveredCellIndex(colIndex)}
                                onMouseLeaveCell={() => setHoveredCellIndex(null)}
                              />
                          ))}
                        </div>
                      </div>
                    );
                  });
                })()}
              </div>
            </div>
          </div>
        )}
      </div>

      {/* 画像拡大モーダル */}
      {zoomedImageIndex !== null && previewImages[zoomedImageIndex] && (
        <div
          onClick={() => setZoomedImageIndex(null)}
          style={{
            position: "fixed",
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            backgroundColor: "rgba(0, 0, 0, 0.9)",
            zIndex: 1000,
            cursor: "zoom-out",
            overflow: "auto",
            WebkitOverflowScrolling: "touch"
          }}
        >
          <div
            onClick={(e) => e.stopPropagation()}
            style={{
              position: "relative",
              display: "inline-block",
              padding: "20px",
              minWidth: "100%",
              minHeight: "100%"
            }}
          >
            <img
              src={previewImages[zoomedImageIndex].imageUrl}
              alt="preview zoomed"
              style={{
                width: "auto",
                height: "auto",
                maxWidth: "200%",
                maxHeight: "none",
                display: "block",
                borderRadius: "8px",
                boxShadow: "0 4px 20px rgba(0,0,0,0.5)",
                cursor: "default"
              }}
            />
            <button
              onClick={() => setZoomedImageIndex(null)}
              style={{
                position: "fixed",
                top: "20px",
                right: "20px",
                backgroundColor: "rgba(255, 255, 255, 0.9)",
                border: "none",
                borderRadius: "50%",
                width: "40px",
                height: "40px",
                fontSize: "24px",
                cursor: "pointer",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                boxShadow: "0 2px 8px rgba(0,0,0,0.3)",
                zIndex: 1001
              }}
            >
              ×
            </button>
            <div style={{
              position: "fixed",
              bottom: "20px",
              left: "50%",
              transform: "translateX(-50%)",
              backgroundColor: "rgba(255, 255, 255, 0.9)",
              padding: "8px 16px",
              borderRadius: "20px",
              fontSize: "14px",
              color: "#111827",
              boxShadow: "0 2px 8px rgba(0,0,0,0.3)"
            }}>
              画像 {zoomedImageIndex + 1} / {previewImages.length}
            </div>
          </div>
        </div>
      )}

      {/* 一括編集モーダル */}
      {bulkModalOpen && (
        <div style={{
          position: "fixed", top: 0, left: 0, right: 0, bottom: 0,
          backgroundColor: "rgba(0,0,0,0.5)", zIndex: 2000,
          display: "flex", alignItems: "center", justifyContent: "center"
        }} onClick={() => setBulkModalOpen(false)}>
          <div style={{
            backgroundColor: "white", padding: "24px", borderRadius: "12px",
            width: "400px", maxWidth: "90%", boxShadow: "0 4px 12px rgba(0,0,0,0.2)"
          }} onClick={e => e.stopPropagation()}>
            <h3 style={{ fontSize: "18px", fontWeight: "bold", marginBottom: "16px" }}>
              {bulkTargetColIndex !== null ? columnHeaders[bulkTargetColIndex] : ""}列の一括編集
            </h3>

            <div style={{ marginBottom: "20px" }}>
              <div style={{ fontWeight: "600", marginBottom: "8px" }}>適用範囲:</div>
              <label style={{ display: "block", marginBottom: "6px" }}>
                <input type="radio" checked={bulkScope === "all"} onChange={() => setBulkScope("all")} /> 全て
              </label>
              <label style={{ display: "block", marginBottom: "6px" }}>
                <input type="radio" checked={bulkScope === "image"} onChange={() => setBulkScope("image")} /> 画像を選択
              </label>
              {bulkScope === "image" && (
                <select
                  style={{ marginLeft: "20px", width: "calc(100% - 20px)", padding: "4px", marginBottom: "6px" }}
                  value={bulkTargetImageIndex}
                  onChange={(e) => setBulkTargetImageIndex(Number(e.target.value))}
                >
                  {previewImages.map((img, idx) => (
                    <option key={idx} value={idx}>{img.fileName} (Image {idx + 1})</option>
                  ))}
                </select>
              )}
              <label style={{ display: "block", marginBottom: "6px" }}>
                <input type="radio" checked={bulkScope === "page"} onChange={() => setBulkScope("page")} /> ページ番号を指定
              </label>
              {bulkScope === "page" && (
                <input
                  type="number" min={1}
                  style={{ marginLeft: "20px", padding: "4px", width: "80px" }}
                  value={bulkTargetPage}
                  onChange={(e) => setBulkTargetPage(Number(e.target.value))}
                />
              )}
            </div>

            <div style={{ display: "flex", flexDirection: "column", gap: "10px" }}>
              <button
                style={{ padding: "10px", backgroundColor: "#3b82f6", color: "white", border: "none", borderRadius: "6px", fontWeight: "bold", cursor: "pointer" }}
                onClick={() => executeBulkUpdate("ok")}
              >
                全て「〇」にする
              </button>
              <button
                style={{ padding: "10px", backgroundColor: "#ef4444", color: "white", border: "none", borderRadius: "6px", fontWeight: "bold", cursor: "pointer" }}
                onClick={() => executeBulkUpdate("ng")}
              >
                全て「×」にする
              </button>
              <button
                style={{ padding: "10px", backgroundColor: "#e5e7eb", color: "#374151", border: "none", borderRadius: "6px", fontWeight: "bold", cursor: "pointer" }}
                onClick={() => setBulkModalOpen(false)}
              >
                キャンセル
              </button>
            </div>
          </div>
        </div>
      )}
    </main>
  );
}

// --- PDF to Images Conversion ---
async function convertPdfToImages(file: File): Promise<PageImage[]> {
  // Dynamic import to avoid SSR issues
  const pdfjsLib = await import('pdfjs-dist');

  // Configure worker (CRITICAL for pdf.js 5.x)
  // Use unpkg CDN which is more reliable
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    `https://unpkg.com/pdfjs-dist@${pdfjsLib.version}/build/pdf.worker.min.mjs`;

  // Load PDF
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;

  const pageImages: PageImage[] = [];
  const numPages = pdf.numPages;

  // Process each page
  for (let pageNum = 1; pageNum <= numPages; pageNum++) {
    const page = await pdf.getPage(pageNum);

    // Set scale for high quality (2x = 144 DPI)
    const scale = 2.0;
    const viewport = page.getViewport({ scale });

    // Create canvas
    const canvas = document.createElement('canvas');
    const context = canvas.getContext('2d');
    if (!context) throw new Error('Canvas context unavailable');

    canvas.width = viewport.width;
    canvas.height = viewport.height;

    // Render page to canvas
    await page.render({
      canvasContext: context,
      viewport: viewport,
      canvas: canvas
    }).promise;

    // Convert canvas to blob (PNG for quality)
    const blob = await new Promise<Blob>((resolve, reject) => {
      canvas.toBlob(
        (b) => b ? resolve(b) : reject(new Error('Blob creation failed')),
        'image/png',
        0.95
      );
    });

    // Create object URL for cell cropping
    const imageUrl = URL.createObjectURL(blob);

    pageImages.push({
      blob,
      imageUrl,
      pageNumber: pageNum,
      width: viewport.width,
      height: viewport.height,
      rotation: 0,
      fileName: `page-${pageNum}`
    });

    // Clean up
    page.cleanup();
  }

  return pageImages;
}

// --- Logic functions (変更なし) ---
async function buildDisplayRows(
  table: Table,
  imageUrl: string,
  imageIndex: number = 0,
  imageName: string = "",
  groupId: number = 0,
  rotatedBlob?: Blob,
  debugMode: boolean = false,
  fallbackStructure?: { indices: number[], headers: string[] }
): Promise<{ displayRows: DisplayRow[], indices: number[], headers: string[], xPositions: number[], prices: Record<string, number> }> {
  // 元のbuildDisplayRows関数（変更なし）
  const rowMap: Record<number, Record<number, string>> = {};

  const filteredCellsGroupedByRow: Record<number, { rowIndex: number; columnIndex: number; polygon: number[]; result: string | null }[]> = {};

  // 1. ヘッダー行と「氏名」列、「合計」列を探す
  let nameRowIndex = -1;
  let nameColumnIndex = -1;
  let totalColumnIndex = -1;
  let shijitsuColumnIndex = -1;

  // デバッグ: テーブルの構造概要のみ出力（詳細はコメントアウト）
  if (debugMode) {
    console.log(`  📋 テーブル構造: ${table.rowCount}行 x ${table.columnCount}列 (${table.cells.length}セル)`);
  }

  // まず「氏名」または「おなまえ」を探す
  for (const cell of table.cells) {
    const content = cell.content?.trim() || "";
    if (content === "氏名" || content.includes("氏名") || content.includes("おなまえ")) {
      nameRowIndex = cell.rowIndex;
      nameColumnIndex = cell.columnIndex;
      if (debugMode) {
        console.log(`  ✓ 氏名/おなまえ列検出: Row=${nameRowIndex}, Col=${nameColumnIndex}, Content="${content}"`);
      }
      break;
    }
  }

  // フォールバックモード: 氏名列なし & 参照ヘッダーあり
  const fallbackMode = (nameRowIndex === -1 || nameColumnIndex === -1) && !!fallbackStructure && fallbackStructure.indices.length > 0;

  if (nameRowIndex === -1 || nameColumnIndex === -1) {
    if (!fallbackMode) {
      console.error("氏名列が見つかりませんでした");
      return { displayRows: [], indices: [], headers: [], xPositions: [], prices: {} };
    }
    console.log("  ⚠️ 氏名列なし → フォールバックモード: 参照ヘッダーを使用");
    // 参照から氏名列の実際の列インデックスを取得（headersで「氏名」を探す）
    const nameIdx = fallbackStructure!.headers.indexOf("氏名");
    nameColumnIndex = nameIdx >= 0 ? fallbackStructure!.indices[nameIdx] : fallbackStructure!.indices[0];
    nameRowIndex = -4; // > -4+3=-1 → 全行(0以上)がデータ行対象
  }

  // 列検出ブロック外で宣言（フォールバックモードでも参照できるよう）
  let sortedXPositions = new Map<number, number>();
  let colorColumnIndices: number[] = [];
  let mergeColorData = false;

  // 同じ行で「合計」などを探す
  for (const cell of table.cells) {
    if (cell.rowIndex !== nameRowIndex) continue;
    const content = cell.content?.trim() || "";

    if (content.includes("合計") || content.includes("小計") || content.includes("金額")) {
      totalColumnIndex = cell.columnIndex;
      if (debugMode) {
        console.log(`  ✓ 合計列検出: Col=${totalColumnIndex}, Content="${content}"`);
      }
    }
    if (content.includes("施術実施") || content.includes("施術有無")) {
      shijitsuColumnIndex = cell.columnIndex;
      if (debugMode) {
        console.log(`  ✓ 施術実施列検出: Col=${shijitsuColumnIndex}, Content="${content}"`);
      }
    }
  }

  // デバッグログは最小限に抑える（パフォーマンス優先）

  // メニュー列の走査範囲: 氏名の次〜合計の手前まで（顔そりは合計の前に必ずある）
  const maxCol = Math.max(...table.cells.map(c => c.columnIndex));
  const searchEndCol = totalColumnIndex !== -1 ? totalColumnIndex : maxCol + 1;

  const targetColumnIndices: number[] = [];
  const columnHeaders: string[] = [];

  // 2. カラム定義の構築 (氏名 -> [メニュー...] -> 施術実施)

  // (A) 氏名列
  targetColumnIndices.push(nameColumnIndex);
  columnHeaders.push("氏名");

  // (B) メニュー列 (氏名 と 合計/右端 の間の列)
  // 単純な行フィルタではなく、列インデックスベースでスキャンする（列の欠落を防ぐため）
  const excludeWords = [
    "備考", "性別", "男女", "メニュー/料金", "メニュー／料金", "メニュー",
    "時間", "料金", "金額", "合計", "合計金額", "小計", "サービス", "項目",
    "内容", "明細", "詳細", "時", "分", "秒", "有無",
    "施術実施", "施術有無", "施術開始時間", "施術開始時間の希望", "ご案内", "案内", "希望",
    "unselected", ":unselected:", "selected", ":selected:",
    "太郎", "山田"  // 例示行「山田 太郎」がセルごとに分かれたときにカラムにならないよう除外
  ];

  // ヘッダーがメニュー項目かどうかを判定するヘルパー関数
  const isValidMenuHeader = (text: string): boolean => {
    if (!text) return false;

    // ※から始まるものはメニューではない（注釈・補足）
    if (text.startsWith("※")) return false;

    // selected/unselected 等の選択マーク状態を除外（大文字小文字区別なし）
    const lower = text.toLowerCase();
    if (lower.includes("unselected") || lower.includes("selected")) return false;

    // 除外ワードに完全一致（トリム後）
    const trimmed = text.trim();
    if (excludeWords.some((w) => trimmed === w || trimmed.toLowerCase() === w.toLowerCase())) return false;

    // 「メニュー」「施術開始」「希望」「案内」を含む文字列を除外
    if (text.includes("メニュー")) return false;
    if (text.includes("施術開始")) return false;
    if (text.includes("希望")) return false;
    if (text.includes("案内")) return false;

    // 数字のみ、または時間形式（例: 10:00, 15分）をフィルタリング
    if (/^[\d:分秒]+$/.test(text)) return false;

    // 時刻・時間表記を除外（9時、30分、9:00、10時30分など）
    if (/^\d+[時分秒]$/.test(text)) return false;  // 9時、30分、45秒
    if (/^\d+時\d*分?$/.test(text)) return false;  // 9時30分、10時
    if (/^\d{1,2}:\d{2}$/.test(text)) return false;  // 9:00、10:30

    // 価格形式（例: ¥1000, 1,000円）をフィルタリング
    if (/^[¥￥$]?[\d,]+円?$/.test(text)) return false;

    // 記号のみをフィルタリング
    if (/^[○×〇✓✗\s]+$/.test(text)) return false;

    // 非常に短い文字列（1文字）をフィルタリング（施術名は通常2文字以上）
    if (text.length < 2) return false;

    // カッコだけの文字列をフィルタリング（例: (注)、【】など）
    if (/^[（）\(\)\[\]【】「」『』]+$/.test(text)) return false;

    // 明らかに誤認識されたメニュー名を除外
    // 「バーマ」は「パーマ」の誤認識なのでフィルタ
    const invalidMenus = ["バーマ", "バ一マ", "ハーマ", "ハ一マ"];
    if (invalidMenus.includes(text)) return false;

    // カタカナで構成されているが長音記号が含まれ、かつ一般的なメニュー名でない場合は除外
    const validKatakanaMenus = [
      "カット", "カラー", "パーマ", "シャンプー", "トリートメント",
      "ヘッドスパ", "ヘアーマニキュア", "パッチテスト", "ベッドカット",
      "ペットカット", "ベットカット"
    ];
    if (/^[ァ-ヴー]+$/.test(text) && text.includes("ー") && !validKatakanaMenus.includes(text)) {
      // 3文字以下の場合は厳しくチェック
      if (text.length <= 3) return false;
    }

    return true;
  };

  // ヘッダー候補: 氏名行の前後3行（マージ・通常候補は±2のみ使い、±3は顔そりフォールバック用）
  const headerCandidates = table.cells.filter(
    (cell) => Math.abs(cell.rowIndex - nameRowIndex) <= 3 && cell.content?.trim()
  );
  const mergedHeaderAssignments: Record<number, string> = {};

  // メニューは指名（氏名）より必ず右側にある → 氏名列より右のセルのみヘッダー候補にする
  const isRightOfNameColumn = (cell: { columnIndex: number }) => cell.columnIndex > nameColumnIndex;

  // 複数メニューが1つのセルにまとまっている場合の処理（±2行のみ、かつ氏名より右の列のみ）
  const headerCandidatesWithin2 = headerCandidates.filter(
    (cell) => Math.abs(cell.rowIndex - nameRowIndex) <= 2 && isRightOfNameColumn(cell)
  );
  headerCandidatesWithin2.forEach((cell) => {
    const span = cell.columnSpan ?? 1;
    const text = cell.content?.trim() || "";

    // 空の場合はスキップ
    if (!text) return;

    // 単一ワードで除外ワードまたは無効なヘッダーの場合はスキップ
    if (!text.includes(' ') && !text.includes('　') && !text.includes('\n')) {
      if (excludeWords.includes(text) || !isValidMenuHeader(text)) {
        return;
      }
    }

    if (debugMode) {
      console.log(`  🔍 ヘッダーセル分析 [${cell.rowIndex},${cell.columnIndex}] span=${span}: "${text}"`);
    }

    // 1. 価格パターンを除去: ¥1,800 や ¥3,800 など
    let cleaned = text.replace(/[¥￥]\s*[\d,]+円?/g, ' ');

    // 2. カッコとその内容を処理
    // 例: "パッチテスト (白髪染め)" → "パッチテスト" と "白髪染め" に分割
    cleaned = cleaned.replace(/\(([^)]+)\)/g, ' $1 ');
    cleaned = cleaned.replace(/\uFF08([^\uFF09]+)\uFF09/g, ' $1 '); // （）
    cleaned = cleaned.replace(/\[([^\]]+)\]/g, ' $1 ');
    cleaned = cleaned.replace(/\u3010([^\u3011]+)\u3011/g, ' $1 '); // 【】

    // 3. スラッシュを区切り文字として扱う
    // 例: "メニュー/料金" → "メニュー" と "料金" に分割
    cleaned = cleaned.replace(/\//g, ' ');

    // 4. スペース、全角スペース、改行、タブなどで分割
    const rawParts = cleaned.split(/[\s\u3000\n\r\t|]+/).filter(Boolean);

    if (debugMode) {
      console.log(`    → 分割前: ${rawParts.length}個`, rawParts);
    }

    // 各パーツを検証してフィルタリング
    const parts = rawParts.filter(part => {
      // 空文字チェック
      if (!part) return false;
      // ※から始まるものはメニューではない
      if (part.startsWith("※")) {
        if (debugMode) console.log(`      ❌ 除外: "${part}" (※注釈)`);
        return false;
      }
      // 除外ワードチェック
      if (excludeWords.includes(part)) {
        if (debugMode) console.log(`      ❌ 除外: "${part}" (除外ワード)`);
        return false;
      }
      // 価格形式チェック
      if (/^[¥￥]/.test(part)) {
        if (debugMode) console.log(`      ❌ 除外: "${part}" (価格)`);
        return false;
      }
      // メニュー項目として有効かチェック
      if (!isValidMenuHeader(part)) {
        if (debugMode) console.log(`      ❌ 除外: "${part}" (無効なヘッダー)`);
        return false;
      }
      if (debugMode) console.log(`      ✅ 有効: "${part}"`);
      return true;
    });

    if (debugMode) {
      console.log(`    → フィルタ後: ${parts.length}個`, parts);
    }

    // パーツがある場合のみ割り当て。その列で始まるセルで既に設定されていれば左から span しているセルで上書きしない
    if (parts.length > 0) {
      for (let offset = 0; offset < parts.length; offset++) {
        const targetCol = cell.columnIndex + offset;
        const isNative = targetCol === cell.columnIndex;
        if (mergedHeaderAssignments[targetCol] === undefined || isNative) {
          mergedHeaderAssignments[targetCol] = parts[offset];
          if (debugMode) {
            console.log(`    ✓ 割り当て: Col ${targetCol} = "${parts[offset]}"${isNative ? " (この列で開始)" : ""}`);
          }
        }
      }
    }
  });

  // 価格情報を抽出（ヘッダー行の近くから）
  const menuPrices: Record<string, number> = {};

  // ヘッダー行の下1-3行以内で価格を探す
  for (let rowOffset = 1; rowOffset <= 3; rowOffset++) {
    const priceRowIndex = nameRowIndex + rowOffset;
    const priceCells = table.cells.filter(cell => cell.rowIndex === priceRowIndex);

    // この行に価格らしきセルがあるかチェック
    const hasPricePattern = priceCells.some(cell => {
      const content = cell.content?.trim() || "";
      return /[¥￥]\s*[\d,]+/.test(content) || /^[\d,]+円?$/.test(content);
    });

    if (hasPricePattern) {
      if (debugMode) {
        console.log(`  💰 価格行を検出: Row ${priceRowIndex}`);
      }

      // 各列の価格を抽出
      priceCells.forEach(cell => {
        const content = cell.content?.trim() || "";
        const priceMatch = content.match(/[¥￥]?\s*([\d,]+)円?/);

        if (priceMatch) {
          const priceValue = parseInt(priceMatch[1].replace(/,/g, ''), 10);

          // この列に対応するメニュー名をmergedHeaderAssignmentsから取得
          const menuName = mergedHeaderAssignments[cell.columnIndex];

          if (menuName && isValidMenuHeader(menuName)) {
            menuPrices[menuName] = priceValue;
            if (debugMode) {
              console.log(`    💴 Col ${cell.columnIndex}: ${menuName} = ¥${priceValue}`);
            }
          } else if (debugMode && content) {
            console.log(`    ⚠️ Col ${cell.columnIndex}: 価格¥${priceValue}があるが、メニュー名が見つからない`);
          }
        }
      });

      break; // 価格行を見つけたらループを抜ける
    }
  }

  if (debugMode && Object.keys(mergedHeaderAssignments).length > 0) {
    console.log(`  📋 マージヘッダー割り当て:`, mergedHeaderAssignments);
  }

  if (debugMode) {
    console.log(`  📊 列スキャン開始: Col ${nameColumnIndex + 1} ~ ${searchEndCol - 1}`);
    if (Object.keys(menuPrices).length > 0) {
      console.log(`  💰 抽出された価格情報:`, menuPrices);
    } else {
      console.log(`  ⚠️ 価格情報が見つかりませんでした`);
    }
  }

  // ヘッダーが "selected" の列は表示しないが、その列のセル値はカラー列の〇/×にマージする
  let selectedColumnIndexForColor: number | null = null;

  for (let c = nameColumnIndex + 1; c < searchEndCol; c++) {
    if (c === shijitsuColumnIndex) continue;

    let headerName = "";

    // この列をカバーするセル（氏名列より右のみ）
    const candidates = headerCandidates.filter((cell) => {
      if (!isRightOfNameColumn(cell)) return false;
      const span = cell.columnSpan ?? 1;
      return cell.columnIndex <= c && c < cell.columnIndex + span;
    });

    // 優先順位1: この列で始まるセルから取得（span で入ったマージ結果より優先して顔そり等を正しく採用）
    const validCandidates = candidates
      .filter(cell => Math.abs(cell.rowIndex - nameRowIndex) <= 2)
      .filter(cell => {
        const text = cell.content?.trim() || "";
        if (text.includes(' ') || text.includes('　') || text.includes('\n') || text.includes('/')) return false;
        return isValidMenuHeader(text);
      })
      .sort((a, b) => {
        const aExact = a.columnIndex === c ? 0 : 1;
        const bExact = b.columnIndex === c ? 0 : 1;
        if (aExact !== bExact) return aExact - bExact;
        return Math.abs(a.rowIndex - nameRowIndex) - Math.abs(b.rowIndex - nameRowIndex);
      });

    if (validCandidates.length > 0) {
      const candidateText = validCandidates[0].content?.trim() || "";
      if (isValidMenuHeader(candidateText)) {
        headerName = candidateText;
        if (debugMode) {
          console.log(`    Col ${c}: 直接セルから = "${headerName}"`);
        }
      }
    }
    // スペース・価格付きセル（例: 「カット ¥1,800」「顔そり ¥500」）から既知メニューを検出（この列で始まるセルを優先）
    if (!headerName && candidates.length > 0) {
      const sortedByRow = [...candidates].sort((a, b) => {
        const aExact = a.columnIndex === c ? 0 : 1;
        const bExact = b.columnIndex === c ? 0 : 1;
        if (aExact !== bExact) return aExact - bExact;
        return Math.abs(a.rowIndex - nameRowIndex) - Math.abs(b.rowIndex - nameRowIndex);
      });
      // 既知メニューのリスト（表記ゆれを含む）
      const knownMenuEntries: [string, string][] = [
        ...Object.keys(DEFAULT_MENU_PRICES).map(name => [name, name] as [string, string]),
        ["顔剃り", "顔そり"],
        ["白髪染め", "カラー"],
        ["ヘアマニキュア", "ヘアーマニキュア"],
        ["ベッドカット", "ベットカット"],
      ];
      for (const cell of sortedByRow) {
        const raw = (cell.content ?? "").trim();
        // 価格・カッコを除去し、スペースを完全除去してノーマライズ（「顔 そり」→「顔そり」対応）
        const cleaned = raw
          .replace(/[¥￥]\s*[\d,]+円?/g, "")
          .replace(/[\d,]+円/g, "")
          .replace(/[（）\(\)\[\]【】「」『』]/g, " ")
          .replace(/\s+/g, "")
          .trim();
        for (const [menuName, canonicalName] of knownMenuEntries) {
          const normalizedMenu = menuName.replace(/\s+/g, "");
          if (cleaned === normalizedMenu) {
            headerName = canonicalName;
            if (debugMode) {
              console.log(`    Col ${c}: 価格付きセルから既知メニュー = "${headerName}" (元: "${raw.slice(0, 30)}")`);
            }
            break;
          }
        }
        if (headerName) break;
      }
    }
    // 優先順位2: まだなければマージされたヘッダー割り当て（span で入った候補）
    if (!headerName && mergedHeaderAssignments[c]) {
      headerName = mergedHeaderAssignments[c];
      if (debugMode) {
        console.log(`    Col ${c}: マージヘッダーから = "${headerName}"`);
      }
    }

    // 選択マーク列 "selected" はヘッダーとしては出さずスキップ。セル値はカラー列の〇/×にマージする
    if (headerName && headerName.trim().toLowerCase().includes("selected")) {
      selectedColumnIndexForColor = c;
      if (debugMode) {
        console.log(`    Col ${c}: "selected" 列をスキップ（カラー列の〇/×にマージするためインデックス記録）`);
      }
      headerName = "";
    }

    // 最終検証（念のため再チェック）
    if (headerName && !isValidMenuHeader(headerName)) {
      if (debugMode) {
        console.log(`    Col ${c}: ❌ 無効なヘッダーを除外 = "${headerName}"`);
      }
      headerName = "";
    }

    // ヘッダー名が無効な場合はスキップ
    if (!headerName) {
      if (debugMode) {
        console.log(`    Col ${c}: ⚠️ スキップ（有効なヘッダーなし）`);
      }
      continue;
    }

    // 白髪染めをカラーに統合
    if (headerName === "白髪染め") {
      console.log(`    Col ${c}: 🔄 "${headerName}" → "カラー" に統合`);
      headerName = "カラー";
    }
    // 顔剃りを顔そりに統合（表記ゆれ対応）
    if (headerName === "顔剃り") {
      console.log(`    Col ${c}: 🔄 "${headerName}" → "顔そり" に統合`);
      headerName = "顔そり";
    }

    targetColumnIndices.push(c);
    columnHeaders.push(headerName);

    if (debugMode) {
      console.log(`    Col ${c}: ✅ 追加 = "${headerName}"`);
    }
  }

  // メニュー列を物理的な位置（X座標）でソート
  // 氏名列のインデックスと名前を保持
  const nameIndex = targetColumnIndices[0];
  const nameHeader = columnHeaders[0];

  // 各列のX座標を取得するヘルパー関数
  const getColumnXPosition = (colIdx: number): number => {
    // この列のセルを探す（ヘッダー行周辺）
    // まず、この列を含む可能性のある全てのセルを探す（結合セルも含む）
    let colCell = table.cells.find(cell =>
      cell.columnIndex === colIdx &&
      Math.abs(cell.rowIndex - nameRowIndex) <= 1
    );

    // 見つからない場合、この列をカバーする結合セルを探す
    if (!colCell) {
      colCell = table.cells.find(cell =>
        cell.columnIndex <= colIdx &&
        colIdx < cell.columnIndex + (cell.columnSpan ?? 1) &&
        Math.abs(cell.rowIndex - nameRowIndex) <= 1
      );
    }

    if (colCell && colCell.boundingRegions && colCell.boundingRegions[0]) {
      const polygon = colCell.boundingRegions[0].polygon;

      // 結合セルの場合、分割後の列位置に応じてX座標を補正
      const span = colCell.columnSpan ?? 1;
      if (span > 1) {
        // 結合セルの幅を分割数で割って、各列の位置を推定
        const leftX = Math.min(polygon[0], polygon[6]);
        const rightX = Math.max(polygon[2], polygon[4]);
        const cellWidth = rightX - leftX;
        const offsetInSpan = colIdx - colCell.columnIndex;
        const estimatedX = leftX + (cellWidth / span) * offsetInSpan;

        if (debugMode) {
          console.log(`    Col ${colIdx}: 結合セルから推定 X=${Math.round(estimatedX)} (span=${span}, offset=${offsetInSpan})`);
        }

        return estimatedX;
      }

      // 通常のセル：左端のX座標を使用
      const leftX = Math.min(polygon[0], polygon[6]);

      if (debugMode) {
        console.log(`    Col ${colIdx}: 直接取得 X=${Math.round(leftX)}`);
      }

      return leftX;
    }

    // 見つからない場合はcolumnIndexをフォールバックとして使用
    if (debugMode) {
      console.log(`    Col ${colIdx}: ⚠️ セルが見つからない、フォールバック X=${colIdx * 100}`);
    }
    return colIdx * 100;
  };

  // 氏名列のX座標を取得
  const nameXPosition = getColumnXPosition(nameIndex);

  // メニュー列（氏名以降）をX座標でソート
  let menuColumns = targetColumnIndices.slice(1).map((colIdx, i) => ({
    index: colIdx,
    header: columnHeaders[i + 1],
    xPosition: getColumnXPosition(colIdx)
  })).sort((a, b) => a.xPosition - b.xPosition);

  // カットとカラーの列が入れ替わっている場合の修正（結合セルでテキスト順が逆の可能性）
  // サロン表では通常 カット | カラー の順（左→右）なので、カットのX座標がカラーより小さいはず
  const cutColumnForSwap = menuColumns.find(col => col.header === "カット");
  const colorColumnForSwap = menuColumns.find(col => col.header === "カラー");
  if (cutColumnForSwap && colorColumnForSwap && cutColumnForSwap.xPosition > colorColumnForSwap.xPosition) {
    console.log(`  🔧 カットとカラーの列を修正: カット(Col${cutColumnForSwap.index}, X=${Math.round(cutColumnForSwap.xPosition)}) が カラー(Col${colorColumnForSwap.index}, X=${Math.round(colorColumnForSwap.xPosition)}) より右に検出 → 列インデックスを入れ替え`);
    const cutOrigIndex = cutColumnForSwap.index;
    const colorOrigIndex = colorColumnForSwap.index;
    cutColumnForSwap.index = colorOrigIndex;
    cutColumnForSwap.xPosition = colorColumnForSwap.xPosition;
    colorColumnForSwap.index = cutOrigIndex;
    colorColumnForSwap.xPosition = getColumnXPosition(cutOrigIndex);
  }

  // カットを強制的に最初に移動
  const cutColumn = menuColumns.find(col => col.header === "カット");
  if (cutColumn) {
    const cutIndex = menuColumns.indexOf(cutColumn);
    if (cutIndex > 0) {
      console.log(`  🔧 カットを先頭に移動: ${cutIndex}番目 → 0番目 (X座標: ${Math.round(cutColumn.xPosition)})`);
      menuColumns = [cutColumn, ...menuColumns.filter(col => col.header !== "カット")];
    }
  }

  // ソート後の配列を再構築
  targetColumnIndices.length = 0;
  columnHeaders.length = 0;

  targetColumnIndices.push(nameIndex);
  columnHeaders.push(nameHeader);

  // ソート済みX座標を保持するマップ
  sortedXPositions = new Map<number, number>();
  sortedXPositions.set(nameIndex, nameXPosition);

  menuColumns.forEach(col => {
    targetColumnIndices.push(col.index);
    columnHeaders.push(col.header);
    sortedXPositions.set(col.index, col.xPosition);
  });

  // カラーの重複を処理（白髪染めがカラーに変換された場合）
  colorColumnIndices = [];
  targetColumnIndices.forEach((colIdx, i) => {
    if (columnHeaders[i] === "カラー") {
      colorColumnIndices.push(colIdx);
    }
  });

  // 重複したカラー列を記録
  mergeColorData = colorColumnIndices.length > 1;
  if (mergeColorData) {
    console.log(`  🔄 カラーが${colorColumnIndices.length}列あります (columnIndex: ${colorColumnIndices.join(', ')})。データを統合して1列にします。`);
  }

  console.log(`  📋 最終的な列構成:`);
  console.log(`     targetColumnIndices: [${targetColumnIndices.join(', ')}]`);
  console.log(`     columnHeaders: [${columnHeaders.join(', ')}]`);

  // カラーヘッダーの重複を削除
  const colorHeaderIndices: number[] = [];
  columnHeaders.forEach((header, i) => {
    if (header === "カラー") {
      colorHeaderIndices.push(i);
    }
  });

  if (colorHeaderIndices.length > 1) {
    // 最初のカラー列だけを残し、他を削除（逆順で削除）
    for (let i = colorHeaderIndices.length - 1; i > 0; i--) {
      const idx = colorHeaderIndices[i];
      console.log(`     ヘッダー削除: ${idx}番目の列 (columnIndex=${targetColumnIndices[idx]})`);
      targetColumnIndices.splice(idx, 1);
      columnHeaders.splice(idx, 1);
    }
  }

  console.log(`  ✅ X座標ソート後のメニュー列: ${columnHeaders.filter(h => h !== '氏名' && h !== '施術実施').join(', ')}`);
  console.log(`     列インデックス: ${targetColumnIndices.join(', ')}`);
  console.log(`     X座標: ${menuColumns.map(c => `${c.header}(X=${Math.round(c.xPosition)})`).join(' → ')}`);

  if (debugMode) {
    console.log(`     詳細: ${menuColumns.map(c => `[${c.header}: idx=${c.index}, X=${Math.round(c.xPosition)}]`).join(', ')}`);
  }

  // (C) 施術実施列を最後に追加
  if (shijitsuColumnIndex !== -1) {
    const shijitsuXPosition = getColumnXPosition(shijitsuColumnIndex);
    targetColumnIndices.push(shijitsuColumnIndex);
    columnHeaders.push("施術実施");
    sortedXPositions.set(shijitsuColumnIndex, shijitsuXPosition);
  }

  // フォールバックモード: 参照ヘッダーで列構成を上書き（位置ベースのマッピング）
  if (fallbackMode && fallbackStructure) {
    console.log("  ⚠️ フォールバック: 参照ヘッダーで列を上書き");
    targetColumnIndices.length = 0;
    columnHeaders.length = 0;
    fallbackStructure.indices.forEach((idx, i) => {
      targetColumnIndices.push(idx);
      columnHeaders.push(fallbackStructure.headers[i]);
    });
    sortedXPositions = new Map<number, number>();
    // colorColumnIndices/mergeColorData を再計算
    colorColumnIndices = [];
    targetColumnIndices.forEach((colIdx, i) => {
      if (columnHeaders[i] === "カラー") colorColumnIndices.push(colIdx);
    });
    mergeColorData = colorColumnIndices.length > 1;
    console.log(`  📋 フォールバック列: [${columnHeaders.join(', ')}]`);
  }

  // Geminiを使ってヘッダーを検証（デバッグモード時のみ）
  // 注: この処理は処理時間が長いため、デバッグモードOFFでは実行されません
  if (debugMode && columnHeaders.length > 1) {
    try {
      console.log(`  🤖 Gemini検証開始: ${columnHeaders.length - 1}個のメニュー候補`);
      const menuCandidates = columnHeaders.filter(h => h !== "氏名" && h !== "施術実施");

      const geminiRes = await fetch('/api/validate-menu', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ menuCandidates }),
      });

      if (geminiRes.ok) {
        const validation = await geminiRes.json();
        console.log(`  ✅ Gemini検証結果:`, validation);

        if (validation.invalidMenus && validation.invalidMenus.length > 0) {
          console.log(`  ⚠️ 無効なメニューを除外: ${validation.invalidMenus.join(', ')}`);

          // 無効なメニューをフィルタリング（顔そり等は必ず残す）
          const alwaysValidMenus = new Set(["顔そり", "顔剃り"]);
          const validIndices: number[] = [];
          const validHeaders: string[] = [];

          validIndices.push(targetColumnIndices[0]);
          validHeaders.push(columnHeaders[0]);

          for (let i = 1; i < columnHeaders.length - 1; i++) {
            if (alwaysValidMenus.has(columnHeaders[i]) || !validation.invalidMenus.includes(columnHeaders[i])) {
              validIndices.push(targetColumnIndices[i]);
              validHeaders.push(columnHeaders[i]);
            }
          }

          if (shijitsuColumnIndex !== -1) {
            validIndices.push(targetColumnIndices[targetColumnIndices.length - 1]);
            validHeaders.push(columnHeaders[columnHeaders.length - 1]);
          }

          targetColumnIndices.length = 0;
          targetColumnIndices.push(...validIndices);
          columnHeaders.length = 0;
          columnHeaders.push(...validHeaders);

          console.log(`  📊 フィルタ後のヘッダー: ${columnHeaders.join(', ')}`);
        }
      } else {
        console.warn('  ⚠️ Gemini検証スキップ: API呼び出し失敗');
      }
    } catch (err) {
      console.warn('  ⚠️ Gemini検証エラー:', err);
    }
  }

  // ログ出力
  if (debugMode) {
    console.log(`  📋 最終ヘッダー: ${columnHeaders.join(', ')}`);
  }

  // データ行の検出
  const targetRowIndices = table.cells
    .filter((c) => c.columnIndex === nameColumnIndex && c.rowIndex > nameRowIndex + 3 && c.content)
    .map((c) => c.rowIndex);

  if (debugMode) {
    console.log(`  📊 データ行検出: ${targetRowIndices.length}行 (Row ${nameRowIndex + 4}以降)`);
  }

  for (const cell of table.cells) {
    if (targetRowIndices.includes(cell.rowIndex) && targetColumnIndices.slice(1).includes(cell.columnIndex) && cell.content) {
      if (!filteredCellsGroupedByRow[cell.rowIndex]) {
        filteredCellsGroupedByRow[cell.rowIndex] = [];
      }
      filteredCellsGroupedByRow[cell.rowIndex].push({
        rowIndex: cell.rowIndex,
        columnIndex: cell.columnIndex,
        polygon: cell.boundingRegions?.[0]?.polygon ?? [],
        result: null,
      });
    }
  }

  // 画像読み込みヘルパー
  const loadImage = (src: string) =>
    new Promise<HTMLImageElement>((resolve, reject) => {
      const img = new Image();
      img.onload = () => resolve(img);
      img.onerror = reject;
      img.src = src;
    });

  let baseImage: HTMLImageElement;
  let objectUrlToRevoke: string | null = null;

  try {
    if (rotatedBlob) {
      objectUrlToRevoke = URL.createObjectURL(rotatedBlob);
      baseImage = await loadImage(objectUrlToRevoke);
    } else {
      baseImage = await loadImage(imageUrl);
    }
  } catch (e) {
    console.error("画像読み込みエラー:", e);
    return { displayRows: [], indices: [], headers: [], xPositions: [], prices: {} };
  }

  const MAX_CONCURRENT = 1;
  let running = 0;
  const queue: (() => Promise<void>)[] = [];

  const enqueue = (task: () => Promise<void>) =>
    new Promise<void>((resolve) => {
      queue.push(async () => {
        running++;
        try { await task(); } finally { running--; resolve(); }
      });
    });

  const runQueue = async () => {
    while (queue.length || running) {
      while (running < MAX_CONCURRENT && queue.length) {
        const job = queue.shift();
        job && job();
      }
      await new Promise((r) => setTimeout(r, 200));
    }
  };

  for (const cells of Object.values(filteredCellsGroupedByRow)) {
    for (const cell of cells) {
      enqueue(async () => {
        const [x1, y1, x2, , , y3] = cell.polygon;
        const w = Math.abs(x2 - x1);
        const h = Math.abs(y3 - y1);
        const canvas = document.createElement("canvas");
        const ctx = canvas.getContext("2d");
        if (!ctx) return;
        canvas.width = w; canvas.height = h;

        ctx.drawImage(baseImage, x1, y1, w, h, 0, 0, w, h);
        const blob = await new Promise<Blob | null>((resolve) => canvas.toBlob(resolve));
        if (!blob) return;

        if (debugMode) {
          const debugUrl = URL.createObjectURL(blob);
          console.log(`🛠️ Debug Crop [Row:${cell.rowIndex}, Col:${cell.columnIndex}]: ${debugUrl}`);
        }

        const fd = new FormData();
        fd.append("image", blob);
        const res = await fetch(`${CUSTOM_VISION_ENDPOINT}customvision/v3.0/Prediction/${PROJECT_ID}/classify/iterations/${ITERATION_ID}/image`, {
          method: "POST", headers: { "Prediction-Key": CUSTOM_VISION_API_KEY }, body: fd,
        });
        if (res.ok) {
          const json = await res.json();
          cell.result = json.predictions?.[0]?.tagName ?? null;
        }
      });
    }
  }
  await runQueue();

  if (objectUrlToRevoke) {
    URL.revokeObjectURL(objectUrlToRevoke);
  }

  // 行範囲内の全列を rowMap に格納（selected 列など表示しない列の〇/×マージ用）
  for (const cell of table.cells) {
    if (cell.rowIndex >= nameRowIndex && cell.rowIndex < (nameRowIndex + targetRowIndices.length + 4)) {
      if (!rowMap[cell.rowIndex]) rowMap[cell.rowIndex] = {};
      rowMap[cell.rowIndex][cell.columnIndex] = cell.content ?? "";
    }
  }

  let yamadaRowIndex: number | null = null;
  let yamadaColumnIndex: number | null = null;
  for (const [r, cols] of Object.entries(rowMap)) {
    for (const [c, value] of Object.entries(cols)) {
      if (value === "山田 太郎") {
        yamadaRowIndex = Number(r);
        yamadaColumnIndex = Number(c);
        break;
      }
    }
    if (yamadaRowIndex !== null) break;
  }

  // ヘッダー行を1つだけ生成
  const headerRow: DisplayRow = {
    rowIndex: nameRowIndex,
    columns: columnHeaders,
    results: columnHeaders.map(() => null),
    sourceImageIndex: imageIndex,
    sourceImageName: imageName,
    groupId: groupId
  };

  // データ行のみを生成（山田行より後の行のみ。山田行なしの場合はヘッダー行より後をすべてデータとして扱う）
  const dataRows = Object.keys(rowMap)
    .map(Number)
    .filter(rowIndex => {
      if (yamadaRowIndex !== null) return rowIndex > yamadaRowIndex;
      if (fallbackMode) return true; // フォールバック+山田行なし → 全行をデータとして扱う
      return rowIndex > nameRowIndex; // 山田行なし → ヘッダー行より後の行をデータとして扱う
    })
    .map((rowIndex, dataRowIdx) => {
      const columns = targetColumnIndices.map((c) => {
        const raw = rowMap[rowIndex]?.[c] ?? "";
        const t = raw.trim().toLowerCase();
        if (t === "selected" || t === "unselected" || t.includes("selected") || t.includes("unselected")) return "";
        return raw;
      });

      // 最初のデータ行の前にヘッダー情報をログ出力
      if (dataRowIdx === 0) {
        console.log(`  🔍 データ行生成デバッグ (Row ${rowIndex}):`);
        console.log(`     columnHeaders: [${columnHeaders.join(', ')}]`);
        console.log(`     targetColumnIndices: [${targetColumnIndices.join(', ')}]`);
      }

      const results = targetColumnIndices.map((c, localIndex) => {
        // 最初のデータ行だけログ出力
        const shouldLog = dataRowIdx === 0;
        // 氏名列（山田列 or nameColumnIndex）より前の列はnull
        const skipUpToColumn = yamadaColumnIndex ?? nameColumnIndex;
        if (c <= skipUpToColumn) {
          return null;
        }

        // カラー列の場合（ヘッダーが"カラー"）、複数のcolumnIndexと selected 列を統合
        if ((mergeColorData || selectedColumnIndexForColor !== null) && columnHeaders[localIndex] === "カラー") {
          let hasCircle = false;
          for (const colorColIdx of colorColumnIndices) {
            const vc = filteredCellsGroupedByRow[rowIndex]?.find((x) => x.columnIndex === colorColIdx);
            const r = vc?.result;
            const cnt = (rowMap[rowIndex]?.[colorColIdx] ?? "").trim();
            const isPresence = r === "Circle" || r === "Check"
              || /^[〇○✓✔︎レ]$/.test(cnt) || /^(チェック|済|済み)$/.test(cnt)
              || cnt.toLowerCase() === "selected" || cnt.toLowerCase().includes("selected");
            if (isPresence) {
              hasCircle = true;
              break;
            }
          }
          if (!hasCircle && selectedColumnIndexForColor !== null) {
            const selCnt = (rowMap[rowIndex]?.[selectedColumnIndexForColor] ?? "").trim().toLowerCase();
            if (selCnt === "selected" || selCnt.includes("selected")) hasCircle = true;
          }
          if (shouldLog) {
            console.log(`      localIndex=${localIndex}(${columnHeaders[localIndex]}): 統合結果=${hasCircle ? "〇" : "×"}`);
          }
          return hasCircle ? "〇" : "×";
        }

        // 通常の列処理
        const visionCell = filteredCellsGroupedByRow[rowIndex]?.find((x) => x.columnIndex === c);
        const result = visionCell?.result;
        const content = (rowMap[rowIndex]?.[c] ?? "").trim();
        const contentLower = content.toLowerCase();

        // 有（〇）とみなす: Vision の Circle/Check、OCR で ✓〇○レチェック等、Azure の選択状態 "selected"
        const isPresence = result === "Circle" || result === "Check"
          || /^[〇○✓✔︎レ]$/.test(content)
          || /^(チェック|済|済み)$/.test(content)
          || contentLower === "selected" || (contentLower.includes("selected") && !contentLower.includes("unselected"));
        // 無（×）とみなす: Vision の Cross/Slash、空、OCR で ×✗、Azure の "unselected"
        const isAbsence = result === "Cross" || result === "Slash" || content === ""
          || /^[×✗]$/.test(content)
          || contentLower === "unselected" || contentLower.includes("unselected");

        if (shouldLog) {
          console.log(`      localIndex=${localIndex}(${columnHeaders[localIndex]}): columnIndex=${c}, result=${result}, content="${content}", final=${isPresence ? "〇" : isAbsence ? "×" : "null"}`);
        }

        if (isPresence) return "〇";
        if (isAbsence) return "×";
        return null;
      });
      return {
        rowIndex,
        columns,
        results,
        sourceImageIndex: imageIndex,
        sourceImageName: imageName,
        groupId: groupId
      };
    })
    .sort((a, b) => a.rowIndex - b.rowIndex);

  const displayRows = [headerRow, ...dataRows];

  // X座標情報を保持（マージ時に順序を正しく保つため）
  // ソート時に計算したX座標を再利用
  const headerXPositions = targetColumnIndices.map(colIdx => {
    return sortedXPositions.get(colIdx) ?? colIdx * 100;
  });

  if (debugMode) {
    console.log(`  📊 DisplayRows生成完了: ${displayRows.length}行 (ヘッダー1行 + データ${dataRows.length}行, 山田行=${yamadaRowIndex})`);
    console.log(`  📍 X座標情報: ${columnHeaders.map((h, i) => `${h}(${Math.round(headerXPositions[i])})`).join(', ')}`);
    if (Object.keys(menuPrices).length > 0) {
      console.log(`  💰 価格情報: ${Object.entries(menuPrices).map(([m, p]) => `${m}(¥${p})`).join(', ')}`);
    }
  }

  return {
    displayRows,
    indices: targetColumnIndices,
    headers: columnHeaders,
    xPositions: headerXPositions,
    prices: menuPrices
  };
}

// --- UI Components ---
function RowView({
  row,
  onToggle,
  onHeaderDelete,
  onHeaderMove,
  onHeaderBulkEdit,
  onRowDelete,
  onRowClick,
  isSelected,
  prices,
  flashColIndex,
  onMouseEnterRow,
  onMouseLeaveRow,
  onMouseEnterCell,
  onMouseLeaveCell,
}: {
  row: DisplayRow;
  onToggle: (rowId: number, colIndex: number) => void;
  onHeaderDelete: (colIndex: number) => void;
  onHeaderMove: (colIndex: number, direction: "left" | "right") => void;
  onHeaderBulkEdit: (colIndex: number) => void;
  onRowDelete: (rowIndex: number) => void;
  onRowClick: (rowIndex: number) => void;
  isSelected: boolean;
  prices?: Record<string, number>;
  flashColIndex?: number | null;
  onMouseEnterRow?: () => void;
  onMouseLeaveRow?: () => void;
  onMouseEnterCell?: (colIndex: number) => void;
  onMouseLeaveCell?: () => void;
}) {
  const isHeaderRow = row.results.every((r) => r === null);

  const rowBg = isHeaderRow ? "#f3f4f6" : isSelected ? "#e5e7eb" : "#ffffff";
  const cellBg = isHeaderRow ? "#f3f4f6" : isSelected ? "#e5e7eb" : "#ffffff";

  return (
    <div
      onMouseEnter={onMouseEnterRow}
      onMouseLeave={onMouseLeaveRow}
      style={{
        display: "flex",
        minWidth: "max-content",
        borderBottom: "1px solid #e5e7eb",
        backgroundColor: rowBg,
        transition: "background-color 0.2s",
        cursor: !isHeaderRow ? "pointer" : "default",
        ...(isHeaderRow ? { position: "sticky" as const, top: 0, zIndex: 2 } : {}),
      }}
      onClick={() => { if (!isHeaderRow) onRowClick(row.rowIndex); }}
    >
      {row.columns.map((c, i) => {
        const result = row.results[i];
        const isName = i === 0;

        const cellRowId = row.rowId ?? row.rowIndex;
        return (
          <div
            key={i}
            data-row-id={isHeaderRow ? undefined : cellRowId}
            data-col-index={i}
            className={flashColIndex === i ? "col-move-flash" : undefined}
            style={{
              width: 120,
              padding: "12px 8px",
              textAlign: "center",
              fontSize: isHeaderRow ? "13px" : "14px",
              fontWeight: isHeaderRow ? "600" : "normal",
              color: isHeaderRow ? "#111827" : "#111827",
              borderRight: "1px solid #e5e7eb",
              display: "flex",
              flexDirection: "column",
              alignItems: "center",
              justifyContent: "center",
              cursor: !isHeaderRow && !isName ? "pointer" : "inherit",
              backgroundColor: cellBg,
              position: "relative",
            }}
            onClick={(e) => {
              if (!isHeaderRow && !isName) {
                e.stopPropagation();
                e.preventDefault();
                const id = e.currentTarget.getAttribute("data-row-id");
                if (id != null) onToggle(Number(id), i);
              }
              // 名前セルの場合は何もせず、イベントを親の行ハンドラーにバブリングさせる
            }}
          >
            {isHeaderRow ? (
              <>
                <div style={{ display: "flex", alignItems: "center", gap: "4px" }}>
                  {c}
                  {c !== "氏名" && (
                    <span
                      role="button"
                      tabIndex={0}
                      onClick={(e) => { e.stopPropagation(); onHeaderBulkEdit(i); }}
                      style={{ fontSize: "12px", cursor: "pointer" }}
                      title="一括編集"
                    >
                      ⚙️
                    </span>
                  )}
                </div>
                {prices && c !== "氏名" && c !== "施術実施" && (
                  <div style={{ fontSize: "11px", color: "#374151", marginTop: "1px", fontWeight: "normal" }}>
                    ¥{(getPriceForHeader(c, prices) ?? 0).toLocaleString()}
                  </div>
                )}
                {isHeaderRow && c !== "氏名" && c !== "施術実施" && (
                  <>
                    <button
                      onClick={(e) => {
                        e.stopPropagation();
                        onHeaderDelete(i);
                      }}
                      style={{
                        position: "absolute", top: "2px", right: "2px", backgroundColor: "#fee2e2",
                        color: "#ef4444", border: "none", borderRadius: "50%", width: "16px", height: "16px",
                        fontSize: "10px", cursor: "pointer", display: "flex", alignItems: "center",
                        justifyContent: "center", padding: 0
                      }}
                      title="列を削除"
                    >
                      ×
                    </button>
                    <div style={{ display: "flex", gap: "4px", marginTop: "6px", justifyContent: "center", flexWrap: "nowrap" }}>
                      <button
                        type="button"
                        onClick={(e) => {
                          e.stopPropagation();
                          e.preventDefault();
                          onHeaderMove(i, "left");
                        }}
                        style={{
                          padding: "4px 8px", fontSize: "11px", border: "1px solid #94a3b8",
                          borderRadius: "4px", backgroundColor: "#e2e8f0", cursor: "pointer",
                          fontWeight: "600"
                        }}
                        title="メニュー順を左へ"
                      >
                        ←
                      </button>
                      <button
                        type="button"
                        onClick={(e) => {
                          e.stopPropagation();
                          e.preventDefault();
                          onHeaderMove(i, "right");
                        }}
                        style={{
                          padding: "4px 8px", fontSize: "11px", border: "1px solid #94a3b8",
                          borderRadius: "4px", backgroundColor: "#e2e8f0", cursor: "pointer",
                          fontWeight: "600"
                        }}
                        title="メニュー順を右へ"
                      >
                        →
                      </button>
                    </div>
                  </>
                )}
              </>
            ) : (
              isName ? (
                <div style={{ position: "relative", width: "100%" }}>
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      onRowDelete(row.rowIndex);
                    }}
                    style={{
                      position: "absolute", left: "-6px", top: "-6px", backgroundColor: "#fecaca",
                      color: "#dc2626", border: "none", borderRadius: "50%", width: "16px", height: "16px",
                      fontSize: "10px", cursor: "pointer", display: "flex", alignItems: "center",
                      justifyContent: "center", padding: 0, zIndex: 10
                    }}
                    title="この人を削除"
                  >
                    ×
                  </button>
                  {c}
                </div>
              ) : (
                <span style={{
                  fontSize: "18px",
                  fontWeight: "bold",
                  color: result === "〇" ? "#dc2626" : "#374151",
                  opacity: 1
                }}>
                  {result === "〇" ? "〇" : "×"}
                </span>
              )
            )}
          </div>
        );
      })}
    </div>
  );
}
