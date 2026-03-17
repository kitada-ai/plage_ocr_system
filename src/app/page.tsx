"use client";

import { useState } from "react";
import InvoiceApp from "../components/InvoiceApp";
import DigitizerApp from "../components/DigitizerApp";

export default function Home() {
  const [activeTab, setActiveTab] = useState<"invoice" | "digitizer">("invoice");

  return (
    <div style={{ display: "flex", flexDirection: "column", minHeight: "100vh", backgroundColor: "#f9fafb" }}>
      {/* 共通のタブナビゲーション */}
      <div style={{ 
        display: "flex", 
        gap: "12px", 
        padding: "16px 24px", 
        backgroundColor: "#111827", 
        boxShadow: "0 4px 6px -1px rgba(0, 0, 0, 0.1)",
        position: "sticky",
        top: 0,
        zIndex: 50
      }}>
        <button
          onClick={() => setActiveTab("invoice")}
          style={{
            padding: "12px 24px",
            backgroundColor: activeTab === "invoice" ? "#2563eb" : "transparent",
            color: activeTab === "invoice" ? "white" : "#9ca3af",
            border: activeTab === "invoice" ? "none" : "1px solid #374151",
            borderRadius: "8px",
            cursor: "pointer",
            fontWeight: "bold",
            fontSize: "16px",
            transition: "all 0.2s",
            display: "flex",
            alignItems: "center",
            gap: "8px"
          }}
        >
          📄 請求書作成システム
        </button>
        <button
          onClick={() => setActiveTab("digitizer")}
          style={{
            padding: "12px 24px",
            backgroundColor: activeTab === "digitizer" ? "#059669" : "transparent",
            color: activeTab === "digitizer" ? "white" : "#9ca3af",
            border: activeTab === "digitizer" ? "none" : "1px solid #374151",
            borderRadius: "8px",
            cursor: "pointer",
            fontWeight: "bold",
            fontSize: "16px",
            transition: "all 0.2s",
            display: "flex",
            alignItems: "center",
            gap: "8px"
          }}
        >
          📝 申込書清書システム
        </button>
      </div>

      {/* コンテンツエリア (マウント状態を維持するために display: none で隠す) */}
      <div style={{ display: activeTab === "invoice" ? "block" : "none", flex: 1 }}>
        <InvoiceApp />
      </div>
      <div style={{ display: activeTab === "digitizer" ? "block" : "none", flex: 1 }}>
        <DigitizerApp />
      </div>
    </div>
  );
}
