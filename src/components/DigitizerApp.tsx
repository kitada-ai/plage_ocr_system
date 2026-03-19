"use client";

import { useState, useEffect, useRef } from "react";

// --- Types ---
type PageImage = {
  blob: Blob;
  imageUrl: string;
  pageNumber: number;
  width: number;
  height: number;
  rotation: number;
  fileName: string;
};

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
  isGuided?: string;
  isAdditionalMenuAllowed?: string;
  isCustomOrder?: string;
  remarks: string;
}

interface ExtractedData {
  facilityName: string;
  year: string;
  month: string;
  day: string;
  dayOfWeek: string;
  menuItems: MenuItem[];
  customers: CustomerData[];
}

export default function DigitizerApp() {
  const [loading, setLoading] = useState(false);
  const [previewImages, setPreviewImages] = useState<PageImage[]>([]);
  const [processingProgress, setProcessingProgress] = useState<{ current: number; total: number }>();
  const [imageZoomLevel, setImageZoomLevel] = useState(100);
  const [zoomedImageIndex, setZoomedImageIndex] = useState<number | null>(null);
  const [statusMessage, setStatusMessage] = useState<string>("");

  // 抽出されたデータ
  const [extractedData, setExtractedData] = useState<ExtractedData | null>(null);
  // 生成されたExcelのURL
  const [excelBlobUrl, setExcelBlobUrl] = useState<string | null>(null);
  const [excelFileName, setExcelFileName] = useState<string>("");

  // メモリクリーンアップ
  useEffect(() => {
    return () => {
      previewImages.forEach((page) => URL.revokeObjectURL(page.imageUrl));
      if (excelBlobUrl) URL.revokeObjectURL(excelBlobUrl);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // --- 画像回転 ---
  const rotateImage = (index: number) => {
    setPreviewImages((prev) => {
      const newImages = [...prev];
      newImages[index] = {
        ...newImages[index],
        rotation: (newImages[index].rotation + 90) % 360,
      };
      return newImages;
    });
  };

  // --- 回転を反映した画像を生成 ---
  const createRotatedImage = async (pageImage: PageImage): Promise<Blob> => {
    if (pageImage.rotation === 0) return pageImage.blob;

    const img = await new Promise<HTMLImageElement>((resolve, reject) => {
      const image = new Image();
      image.onload = () => resolve(image);
      image.onerror = reject;
      image.src = pageImage.imageUrl;
    });

    const canvas = document.createElement("canvas");
    const ctx = canvas.getContext("2d");
    if (!ctx) throw new Error("Canvas context not available");

    const rotation = pageImage.rotation;
    if (rotation === 90 || rotation === 270) {
      canvas.width = img.height;
      canvas.height = img.width;
    } else {
      canvas.width = img.width;
      canvas.height = img.height;
    }

    ctx.translate(canvas.width / 2, canvas.height / 2);
    ctx.rotate((rotation * Math.PI) / 180);
    ctx.drawImage(img, -img.width / 2, -img.height / 2);

    return new Promise((resolve, reject) => {
      canvas.toBlob(
        (blob) => (blob ? resolve(blob) : reject(new Error("Blob creation failed"))),
        "image/png",
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

  // --- Excelファイルを生成してダウンロード ---
  const generateAndDownloadExcel = async (data: ExtractedData) => {
    const payload = {
      facilityName: data.facilityName,
      year: data.year,
      month: data.month,
      day: data.day,
      dayOfWeek: data.dayOfWeek,
      menuItems: data.menuItems,
      customers: data.customers,
    };

    const res = await fetch("/api/export-digitized-excel", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });

    if (!res.ok) {
      const errorText = await res.text();
      throw new Error(`Excel生成失敗: ${errorText}`);
    }

    const blob = await res.blob();
    const url = URL.createObjectURL(blob);

    // 前のURLを解放
    if (excelBlobUrl) URL.revokeObjectURL(excelBlobUrl);

    const fileDate = `${data.year || ""}${data.month || ""}${data.day || ""}`;
    const fileName = `${data.facilityName || "申込書"}_${fileDate || "清書"}.xlsx`;

    setExcelBlobUrl(url);
    setExcelFileName(fileName);

    // 自動ダウンロード（Excelで直接開く）
    const a = document.createElement("a");
    a.href = url;
    a.download = fileName;
    a.click();

    return { url, fileName };
  };

  // --- 解析実行 ---
  const onSubmit = async () => {
    if (previewImages.length === 0) {
      alert("画像を選択してください。");
      return;
    }

    setLoading(true);
    setExtractedData(null);
    setExcelBlobUrl(null);
    setStatusMessage("OCR処理を開始します...");

    try {
      const totalPages = previewImages.length;
      setProcessingProgress({ current: 0, total: totalPages });

      // 全ページのOCRテキストを集約
      let allMarkdown = "";

      for (let i = 0; i < previewImages.length; i++) {
        const pageImage = previewImages[i];
        setStatusMessage(`OCR処理中: ${i + 1}/${previewImages.length} ページ...`);

        const rotatedBlob = await createRotatedImage(pageImage);
        const formData = new FormData();
        formData.append("file", rotatedBlob, `${pageImage.fileName}.png`);

        let res;
        let data;
        let retryCount = 0;
        const maxRetries = 5;

        while (retryCount <= maxRetries) {
          try {
            res = await fetch("/api/analyze", { method: "POST", body: formData });

            if (res.status === 403 || res.status === 429) {
              if (retryCount < maxRetries) {
                const waitTime = (5 + retryCount * 2) * 1000;
                setStatusMessage(`API制限。${waitTime / 1000}秒待機後にリトライ...`);
                await new Promise((r) => setTimeout(r, waitTime));
                retryCount++;
                continue;
              }
              throw new Error(`APIアクセスが拒否されました (${res.status})`);
            }

            if (!res.ok) throw new Error(`API呼び出しが失敗しました: ${res.status}`);
            data = await res.json();
            break;
          } catch (error) {
            if (retryCount >= maxRetries) throw error;
            retryCount++;
            await new Promise((r) => setTimeout(r, 2000));
          }
        }

        const pageMarkdown = data?.analyzeResult?.content || "";
        if (pageMarkdown) {
          allMarkdown += `\n--- ページ ${i + 1} ---\n${pageMarkdown}`;
        }

        setProcessingProgress({ current: i + 1, total: totalPages });

        if (i < previewImages.length - 1) {
          await new Promise((r) => setTimeout(r, 1000));
        }
      }

      if (!allMarkdown.trim()) {
        alert("OCRテキストが取得できませんでした。");
        return;
      }

      // Gemini でフル情報を一括抽出
      setStatusMessage("AIでデータを抽出中...");

      const extractRes = await fetch("/api/extract-full", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ text: allMarkdown }),
      });

      if (!extractRes.ok) {
        const errText = await extractRes.text();
        throw new Error(`Gemini抽出失敗: ${errText}`);
      }

      const extracted: ExtractedData = await extractRes.json();
      console.log("✅ Gemini抽出完了:", extracted);
      setExtractedData(extracted);

      // Excelファイルを生成して自動ダウンロード
      setStatusMessage("Excelファイルを生成中...");
      await generateAndDownloadExcel(extracted);
      setStatusMessage("✅ Excelファイルを生成しました。ブラウザのダウンロード設定を「常に開く」にすると自動でExcelが起動します。");

    } catch (error) {
      console.error("❌ 解析エラー:", error);
      let errorMessage = "解析中にエラーが発生しました。";
      if (error instanceof Error) {
        errorMessage += `\n\nエラー詳細: ${error.message}`;
      }
      setStatusMessage("");
      alert(errorMessage);
    } finally {
      setLoading(false);
      setProcessingProgress(undefined);
    }
  };

  // --- 再ダウンロード ---
  const onRedownload = () => {
    if (excelBlobUrl && excelFileName) {
      const a = document.createElement("a");
      a.href = excelBlobUrl;
      a.download = excelFileName;
      a.click();
    }
  };

  // --- 抽出データ修正後に再生成 ---
  const onRegenerateExcel = async () => {
    if (!extractedData) return;
    setLoading(true);
    try {
      setStatusMessage("Excelファイルを再生成中...");
      await generateAndDownloadExcel(extractedData);
      setStatusMessage("✅ Excelファイルを再生成しました。ブラウザの設定で「常に開く」を選択すると便利です。");
    } catch (err) {
      console.error("Excel再生成エラー:", err);
      alert("Excel再生成中にエラーが発生しました。");
    } finally {
      setLoading(false);
    }
  };

  return (
    <main
      style={{
        width: "100%",
        maxWidth: "100vw",
        margin: 0,
        padding: "24px",
        fontFamily: "sans-serif",
        backgroundColor: "#f9fafb",
        color: "#111827",
        minHeight: "100vh",
        boxSizing: "border-box",
      }}
    >
      {/* ヘッダー */}
      <div
        style={{
          backgroundColor: "white",
          padding: "24px",
          borderRadius: "12px",
          boxShadow: "0 1px 3px rgba(0,0,0,0.1)",
          marginBottom: "24px",
        }}
      >
        <h1 style={{ fontSize: "24px", fontWeight: "bold", marginBottom: "20px" }}>
          申込書清書システム
        </h1>

        <div style={{ display: "flex", flexWrap: "wrap", gap: "12px", alignItems: "center", marginBottom: "16px" }}>
          <input
            type="file"
            accept="image/*,application/pdf"
            multiple
            style={{
              padding: "8px",
              border: "1px solid #d1d5db",
              borderRadius: "6px",
              flex: "1",
              minWidth: "200px",
            }}
            onChange={async (e) => {
              const fileList = e.target.files;
              if (fileList && fileList.length > 0) {
                const filesArray = Array.from(fileList);
                setLoading(true);
                try {
                  const newImages: PageImage[] = [];
                  for (const file of filesArray) {
                    if (file.type === "application/pdf") {
                      const pageImgs = await convertPdfToImages(file);
                      pageImgs.forEach((img) => {
                        img.fileName = `${file.name} - ページ ${img.pageNumber}`;
                      });
                      newImages.push(...pageImgs);
                    } else {
                      const imageUrl = URL.createObjectURL(file);
                      const img = await new Promise<HTMLImageElement>((resolve, reject) => {
                        const image = new Image();
                        image.onload = () => resolve(image);
                        image.onerror = reject;
                        image.src = imageUrl;
                      });
                      newImages.push({
                        blob: file,
                        imageUrl,
                        pageNumber: 1,
                        width: img.width,
                        height: img.height,
                        rotation: 0,
                        fileName: file.name,
                      });
                    }
                  }
                  setPreviewImages((prev) => [...prev, ...newImages]);
                } catch (error) {
                  console.error("プレビュー生成エラー:", error);
                  alert("プレビューの生成中にエラーが発生しました。");
                } finally {
                  setLoading(false);
                }
              }
              e.target.value = "";
            }}
          />
          <button
            onClick={onSubmit}
            disabled={loading || previewImages.length === 0}
            style={{
              padding: "10px 24px",
              backgroundColor: loading ? "#9ca3af" : "#2563eb",
              color: "white",
              border: "none",
              borderRadius: "6px",
              cursor: loading ? "not-allowed" : "pointer",
              fontWeight: "600",
              fontSize: "15px",
            }}
          >
            {loading ? "処理中..." : "📤 アップロード & Excel生成"}
          </button>
          {excelBlobUrl && (
            <>
              <button
                onClick={onRedownload}
                style={{
                  padding: "10px 20px",
                  backgroundColor: "#059669",
                  color: "white",
                  border: "none",
                  borderRadius: "6px",
                  cursor: "pointer",
                  fontWeight: "600",
                  display: "flex",
                  alignItems: "center",
                  gap: "6px",
                }}
              >
                📥 再ダウンロード
              </button>
              <button
                onClick={onRegenerateExcel}
                disabled={loading}
                style={{
                  padding: "10px 20px",
                  backgroundColor: "#7c3aed",
                  color: "white",
                  border: "none",
                  borderRadius: "6px",
                  cursor: loading ? "not-allowed" : "pointer",
                  fontWeight: "600",
                }}
              >
                🔄 Excel再生成
              </button>
            </>
          )}
        </div>

        {/* ステータスメッセージ */}
        {statusMessage && (
          <div
            style={{
              padding: "12px 16px",
              backgroundColor: statusMessage.startsWith("✅") ? "#f0fdf4" : "#eff6ff",
              color: statusMessage.startsWith("✅") ? "#166534" : "#1e40af",
              borderRadius: "8px",
              fontSize: "14px",
              fontWeight: "500",
              marginBottom: "16px",
              border: `1px solid ${statusMessage.startsWith("✅") ? "#bbf7d0" : "#bfdbfe"}`,
            }}
          >
            {statusMessage}
          </div>
        )}

        {processingProgress && (
          <div
            style={{
              padding: "8px 16px",
              backgroundColor: "#eff6ff",
              borderRadius: "6px",
              fontSize: "14px",
              color: "#1e40af",
              marginBottom: "16px",
            }}
          >
            処理中: {processingProgress.current} / {processingProgress.total} ページ
          </div>
        )}
      </div>

      {/* メインコンテンツ: 画像プレビュー */}
      <div
        style={{
          display: "flex",
          gap: "24px",
          width: "100%",
          alignItems: "flex-start",
        }}
      >
        {/* 左側: 画像プレビュー */}
        <div
          style={{
            flex: "1",
            backgroundColor: "white",
            padding: "24px",
            borderRadius: "12px",
            boxShadow: "0 1px 3px rgba(0,0,0,0.1)",
            display: "flex",
            flexDirection: "column",
          }}
        >
          <div
            style={{
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
              marginBottom: "16px",
            }}
          >
            <h2 style={{ fontSize: "18px", fontWeight: "bold", margin: 0 }}>
              画像プレビュー {previewImages.length > 0 ? `(${previewImages.length}枚)` : ""}
            </h2>
            <div style={{ display: "flex", gap: "8px", alignItems: "center" }}>
              <button
                onClick={() => setImageZoomLevel((prev) => Math.max(30, prev - 10))}
                style={{
                  padding: "4px 12px",
                  backgroundColor: "#f3f4f6",
                  border: "1px solid #d1d5db",
                  borderRadius: "4px",
                  fontSize: "14px",
                  cursor: "pointer",
                  fontWeight: "bold",
                }}
              >
                −
              </button>
              <span
                style={{
                  padding: "4px 12px",
                  backgroundColor: "#f9fafb",
                  border: "1px solid #e5e7eb",
                  borderRadius: "4px",
                  fontSize: "12px",
                  minWidth: "50px",
                  textAlign: "center",
                }}
              >
                {imageZoomLevel}%
              </span>
              <button
                onClick={() => setImageZoomLevel((prev) => Math.min(300, prev + 10))}
                style={{
                  padding: "4px 12px",
                  backgroundColor: "#f3f4f6",
                  border: "1px solid #d1d5db",
                  borderRadius: "4px",
                  fontSize: "14px",
                  cursor: "pointer",
                  fontWeight: "bold",
                }}
              >
                ＋
              </button>
            </div>
          </div>

          <div style={{ overflowY: "auto", maxHeight: "calc(100vh - 300px)" }}>
            {previewImages.length > 0 ? (
              <div style={{ display: "flex", flexDirection: "column", gap: "16px" }}>
                {previewImages.map((page, index) => (
                  <div
                    key={index}
                    style={{
                      border: "1px solid #e5e7eb",
                      borderRadius: "8px",
                      overflow: "hidden",
                      position: "relative",
                    }}
                  >
                    <button
                      onClick={(e) => { e.stopPropagation(); removePreviewImage(index); }}
                      style={{
                        position: "absolute", top: "8px", left: "8px",
                        backgroundColor: "#fff", border: "1px solid #d1d5db",
                        borderRadius: "4px", padding: "6px 10px", fontSize: "12px",
                        cursor: "pointer", fontWeight: "600", boxShadow: "0 2px 4px rgba(0,0,0,0.1)", zIndex: 10,
                      }}
                    >
                      🗑️ 削除
                    </button>
                    <button
                      onClick={(e) => { e.stopPropagation(); rotateImage(index); }}
                      style={{
                        position: "absolute", top: "8px", right: "8px",
                        backgroundColor: "#fff", border: "1px solid #d1d5db",
                        borderRadius: "4px", padding: "6px 12px", fontSize: "12px",
                        cursor: "pointer", fontWeight: "600", boxShadow: "0 2px 4px rgba(0,0,0,0.1)", zIndex: 10,
                      }}
                    >
                      🔄 回転
                    </button>
                    <div onClick={() => setZoomedImageIndex(index)} style={{ cursor: "zoom-in" }}>
                      <img
                        src={page.imageUrl}
                        alt={`Image ${index + 1}`}
                        style={{
                          height: `${(600 * imageZoomLevel) / 100}px`,
                          width: "auto",
                          display: "block",
                          objectFit: "contain",
                          transform: `rotate(${page.rotation}deg)`,
                          transition: "transform 0.3s ease",
                        }}
                      />
                    </div>
                    <div
                      style={{
                        padding: "8px", backgroundColor: "#f3f4f6",
                        textAlign: "center", fontSize: "11px",
                      }}
                    >
                      <div style={{ fontWeight: "600" }}>{page.fileName}</div>
                      <div style={{ color: "#4b5563", fontSize: "10px" }}>回転: {page.rotation}°</div>
                    </div>
                  </div>
                ))}
              </div>
            ) : (
              <div
                style={{
                  display: "flex", alignItems: "center", justifyContent: "center",
                  height: "200px", color: "#4b5563", fontSize: "14px",
                }}
              >
                画像を選択してください
              </div>
            )}
          </div>
        </div>

        {/* 右側: 抽出結果サマリ */}
        {extractedData && (
          <div
            style={{
              flex: "0 0 400px",
              backgroundColor: "white",
              padding: "24px",
              borderRadius: "12px",
              boxShadow: "0 1px 3px rgba(0,0,0,0.1)",
            }}
          >
            <h2 style={{ fontSize: "18px", fontWeight: "bold", marginBottom: "16px" }}>
              📋 抽出結果
            </h2>

            <div style={{ display: "flex", flexDirection: "column", gap: "12px" }}>
              <div style={{ padding: "12px", backgroundColor: "#f0fdf4", borderRadius: "8px", border: "1px solid #bbf7d0" }}>
                <div style={{ fontWeight: "600", marginBottom: "4px", color: "#166534" }}>施設名</div>
                <div style={{ fontSize: "16px" }}>{extractedData.facilityName || "(未取得)"}</div>
              </div>

              <div style={{ padding: "12px", backgroundColor: "#eff6ff", borderRadius: "8px", border: "1px solid #bfdbfe" }}>
                <div style={{ fontWeight: "600", marginBottom: "4px", color: "#1e40af" }}>施術日</div>
                <div style={{ fontSize: "16px" }}>
                  {extractedData.year && extractedData.month && extractedData.day
                    ? `${extractedData.year}年${extractedData.month}月${extractedData.day}日（${extractedData.dayOfWeek || ""}）`
                    : "(未取得)"}
                </div>
              </div>

              <div style={{ padding: "12px", backgroundColor: "#fefce8", borderRadius: "8px", border: "1px solid #fef08a" }}>
                <div style={{ fontWeight: "600", marginBottom: "8px", color: "#854d0e" }}>メニュー</div>
                {extractedData.menuItems.map((item, i) => (
                  <div key={i} style={{ display: "flex", justifyContent: "space-between", padding: "4px 0", borderBottom: i < extractedData.menuItems.length - 1 ? "1px solid #fef08a" : "none" }}>
                    <span>{item.name}</span>
                    <span style={{ fontWeight: "500" }}>¥{item.price.toLocaleString()}</span>
                  </div>
                ))}
              </div>

              <div style={{ padding: "12px", backgroundColor: "#fdf4ff", borderRadius: "8px", border: "1px solid #f0abfc" }}>
                <div style={{ fontWeight: "600", marginBottom: "8px", color: "#86198f" }}>
                  顧客データ ({extractedData.customers.length}名)
                </div>
                {extractedData.customers.map((c, i) => (
                  <div key={i} style={{ display: "flex", gap: "8px", padding: "4px 0", borderBottom: i < extractedData.customers.length - 1 ? "1px solid #f0abfc" : "none", fontSize: "13px" }}>
                    <span style={{ fontWeight: "500", minWidth: "80px" }}>{c.name}</span>
                    <div style={{ flex: 1 }}>
                      <div style={{ color: "#4b5563" }}>{c.selectedMenus.join(", ") || "なし"}</div>
                      {(c.isGuided || c.isAdditionalMenuAllowed || c.isCustomOrder) && (
                        <div style={{ fontSize: "11px", color: "#9333ea", marginTop: "2px" }}>
                          {[
                            c.isCustomOrder ? `オーダー: ${c.isCustomOrder}` : null,
                            c.isGuided ? `案内: ${c.isGuided}` : null,
                            c.isAdditionalMenuAllowed ? `追加メ: ${c.isAdditionalMenuAllowed}` : null,
                          ].filter(Boolean).join(" / ")}
                        </div>
                      )}
                    </div>
                  </div>
                ))}
              </div>
            </div>

            <div style={{ marginTop: "20px", padding: "16px", backgroundColor: "#f8fafc", borderRadius: "8px", border: "1px solid #e2e8f0" }}>
              <p style={{ fontSize: "13px", color: "#64748b", margin: 0 }}>
                💡 ダウンロードされたExcelファイルを直接開いて編集してください。
                編集後はそのまま保存できます。
              </p>
            </div>
          </div>
        )}
      </div>

      {/* 画像拡大モーダル */}
      {zoomedImageIndex !== null && previewImages[zoomedImageIndex] && (
        <div
          onClick={() => setZoomedImageIndex(null)}
          style={{
            position: "fixed", top: 0, left: 0, right: 0, bottom: 0,
            backgroundColor: "rgba(0,0,0,0.9)", zIndex: 1000,
            cursor: "zoom-out", overflow: "auto",
          }}
        >
          <div
            onClick={(e) => e.stopPropagation()}
            style={{ position: "relative", display: "inline-block", padding: "20px", minWidth: "100%", minHeight: "100%" }}
          >
            <img
              src={previewImages[zoomedImageIndex].imageUrl}
              alt="preview zoomed"
              style={{
                width: "auto", height: "auto", maxWidth: "200%",
                display: "block", borderRadius: "8px",
                boxShadow: "0 4px 20px rgba(0,0,0,0.5)", cursor: "default",
              }}
            />
            <button
              onClick={() => setZoomedImageIndex(null)}
              style={{
                position: "fixed", top: "20px", right: "20px",
                backgroundColor: "rgba(255,255,255,0.9)", border: "none",
                borderRadius: "50%", width: "40px", height: "40px",
                fontSize: "24px", cursor: "pointer",
                display: "flex", alignItems: "center", justifyContent: "center",
                boxShadow: "0 2px 8px rgba(0,0,0,0.3)", zIndex: 1001,
              }}
            >
              ×
            </button>
          </div>
        </div>
      )}
    </main>
  );
}

// --- PDF to Images Conversion ---
async function convertPdfToImages(file: File): Promise<PageImage[]> {
  const pdfjsLib = await import("pdfjs-dist");
  pdfjsLib.GlobalWorkerOptions.workerSrc = `https://unpkg.com/pdfjs-dist@${pdfjsLib.version}/build/pdf.worker.min.mjs`;

  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;

  const pageImages: PageImage[] = [];

  for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
    const page = await pdf.getPage(pageNum);
    const scale = 2.0;
    const viewport = page.getViewport({ scale });

    const canvas = document.createElement("canvas");
    const context = canvas.getContext("2d");
    if (!context) throw new Error("Canvas context unavailable");

    canvas.width = viewport.width;
    canvas.height = viewport.height;

    await page.render({ canvasContext: context, viewport, canvas }).promise;

    const blob = await new Promise<Blob>((resolve, reject) => {
      canvas.toBlob(
        (b) => (b ? resolve(b) : reject(new Error("Blob creation failed"))),
        "image/png",
        0.95
      );
    });

    const imageUrl = URL.createObjectURL(blob);

    pageImages.push({
      blob,
      imageUrl,
      pageNumber: pageNum,
      width: viewport.width,
      height: viewport.height,
      rotation: 0,
      fileName: `page-${pageNum}`,
    });

    page.cleanup();
  }

  return pageImages;
}
