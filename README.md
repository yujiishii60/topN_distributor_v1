🧾 topN_distributor_v1

店舗別 × 大分類別 × 任意日数の TopN単品データ配布ファイル を自動生成するツール。
複数日の売上データを読み込み、テンプレート書式を維持したまま Excel 出力します。

📂 構成
topN_distributor_v1/
├─ data/
│  ├─ material/
│  │  ├─ 2024/IT_202412.csv
│  │  ├─ 2025/IT_202501.csv
│  │  └─ master/store_master.xlsx
│  ├─ template/配布フォーマット.xlsx
│  └─ output/
│     ├─ topN_寿司_年末年始16日.xlsx       ← まとめ版
│     └─ split/
│        ├─ 1/1_寿司単品データ.xlsx
│        ├─ 2/2_寿司単品データ.xlsx
│        └─ …（店別出力）
├─ config/
│  └─ category_map.json
└─ scripts/
   └─ make_topn_simple_refactor.py

🚀 基本の使い方

PowerShell から以下を実行：

```
Set-Location C:\Users\14ugy\Projects\topN_distributor_v1
$env:PYTHONPATH = (Get-Location).Path

python -m scripts.make_topn_simple_refactor `
  --category 1 `
  --dates "2024-12-20,2024-12-21,2024-12-22,2024-12-23,2024-12-24,2024-12-25,2024-12-26,2024-12-27,2024-12-28,2024-12-29,2024-12-30,2024-12-31,2025-01-02,2025-01-03,2025-01-04,2025-01-05" `
  --out data/output/topN_寿司_年末年始16日.xlsx `
  --split-by-store `
  --split-dir data/output/split `
  --event-name "2024-2025年　年末年始" `
  --no-date-in-title
```

✅ 出力内容

data/output/topN_寿司_年末年始16日.xlsx
→ まとめ版（4日ごと×4シート）

data/output/split/<店番>/<店番>_寿司単品データ.xlsx
→ 店別配布版。タイトルは
2024-2025年　年末年始 寿司単品データ (1) の形式で日付なし。

🧩 主なオプション
オプション名	説明
--category	大分類コード（例：1=寿司）
--dates	対象日（カンマ区切り YYYY-MM-DD）
--out	まとめ版Excelの出力パス
--split-by-store	店別にファイル分割
--split-dir	店別ファイルの出力先ルート
--event-name	タイトル先頭のイベント名（例：「2024-2025年　年末年始」）
--no-date-in-title	タイトルから日付を除外（{event} {cat}単品データ ({page})）
--title-template	A1タイトルの完全カスタムテンプレ（例："{event} {cat}配布用 ({page})"）

🗂️ カテゴリマップ設定

カテゴリ名は config/category_map.json で管理：

{
  "1": "寿司",
  "2": "弁当",
  "3": "温総菜",
  "4": "冷総菜",
  "5": "軽食",
  "6": "魚惣菜"
}

存在しない場合や壊れている場合は内蔵マップにフォールバックします。

⚙️ 動作仕様

年月跨ぎ自動読込
指定日の年ごとに data/material/YYYY/IT_YYYYMM.csv を自動選択。

ページ分割
8日以上 → 4日ごとに自動で (1)(2)... のシートを生成。

店別スプリット
--split-by-store 指定で <店番>/<店番>_<カテゴリ名>単品データ.xlsx を自動生成。

テンプレート維持
書式・罫線・条件付き書式・印刷設定を保持。

合計・構成比・値引率
自動計算済み。小数点・桁区切りもテンプレ仕様に合わせて出力。

🧪 CI テスト想定
テスト内容	目的
年跨ぎ実行（2024-12〜2025-01）	複数CSVの結合確認
8日実行	ページ分割(4日×2シート)の確認
店別スプリット結果	出力フォルダ構成とタイトル確認
--no-date-in-title 実行	タイトルから日付除外動作確認

🏁 出力例
タイトル: 2024-2025年　年末年始 寿司単品データ (1)
シート構成: (1)(2)(3)(4)
ファイル: data/output/split/1/1_寿司単品データ.xlsx

依存パッケージ: tkcalendar (pip install tkcalendar)
