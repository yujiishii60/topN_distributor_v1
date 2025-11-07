# scripts/make_topn_simple_refactor.py
import pandas as pd
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import math
from openpyxl.formatting.rule import Rule

CATEGORY_MAP = {
    "1": "寿司", "2": "米飯", "3": "温惣菜",
    "4": "冷総菜", "5": "軽食", "6": "魚惣菜",
}

# 先頭でimport

def copy_conditional_formatting(ws_dst, ws_src):
    """
    copy_worksheet で失われがちな条件付き書式を、TEMPLATE から新シートへ再適用する。
    レイアウトが同一（セル座標が同じ）前提で、そのまま同レンジへ貼る。
    """
    cf_src = ws_src.conditional_formatting
    # openpyxlの内部構造はバージョンで差があります。代表的な2系統に対応。
    if hasattr(cf_src, "cf_rules"):  # 3.1系で公開属性がある場合
        items = cf_src.cf_rules.items()
    else:  # 旧来: _cf_rules にレンジ→ルールlist が入っていることが多い
        items = getattr(cf_src, "_cf_rules", {}).items()

    for rng, rules in items:
        for rule in rules:
            # そのまま add すると同じオブジェクト参照になることがあるので clone 相当を作る
            new_rule = Rule(
                type=rule.type,
                dxf=rule.dxf,
                formula=list(rule.formula) if hasattr(rule, "formula") else None,
                operator=getattr(rule, "operator", None),
                text=getattr(rule, "text", None),
                timePeriod=getattr(rule, "timePeriod", None),
                rank=getattr(rule, "rank", None),
                percent=getattr(rule, "percent", None),
                stopIfTrue=getattr(rule, "stopIfTrue", False),
            )
            # カラースケール/データバーなど複合型も転写
            for attr in ("colorScale", "dataBar", "iconSet"):
                if hasattr(rule, attr) and getattr(rule, attr):
                    setattr(new_rule, attr, getattr(rule, attr))
            ws_dst.conditional_formatting.add(rng, new_rule)

# === CSV読込 ===
def load_sales(csv_root: Path) -> pd.DataFrame:
    files = list(csv_root.glob("2025/*.csv"))
    if not files:
        raise FileNotFoundError(f"no csv files under {csv_root/'2025'}")

    def read_csv_any(path: Path) -> pd.DataFrame:
        for enc in ("cp932", "utf-8-sig", "utf-8"):
            try:
                return pd.read_csv(path, encoding=enc)
            except Exception:
                continue
            df_day = df_day.sort_values("amount", ascending=False)  # ←ここ！
        raise UnicodeDecodeError("utf-8", b"", 0, 1, f"Failed to decode {path}")

    df = pd.concat((read_csv_any(f) for f in files), ignore_index=True)

    # === 列名を英名に正規化 ===
    df = df.rename(columns={
        "売上日": "date",
        "店舗コード": "store_id",
        "大分類コード": "category_large",
        "中分類コード": "category_middle",
        "小分類コード": "category_small",
        "JANコード": "jan",
        "品名漢字": "name",
        "総売上金額": "amount",
        "総売上数量": "qty",
        "値引金額": "discount",
    })

    # === 型変換・整形 ===
    df["date"] = pd.to_datetime(df["date"], errors="coerce").dt.date
    df["store_id"] = df["store_id"].astype(str)
    df["jan"] = df["jan"].astype(str).str.strip()
    df["amount"] = pd.to_numeric(df["amount"], errors="coerce").fillna(0)
    df["qty"] = pd.to_numeric(df["qty"], errors="coerce").fillna(0)
    df["discount"] = pd.to_numeric(df.get("discount", 0), errors="coerce").fillna(0)
    return df

# === 店舗マスター ===
def load_store_master(path: Path) -> dict:
    sm = pd.read_excel(path)
    sm = sm.rename(columns={
        "store": "store_id",
        "name": "store_name",
        "short_name": "short_name",
    })
    sm["store_id"] = sm["store_id"].astype(str)
    return sm.set_index("store_id")["short_name"].to_dict()


# === 大分類・日付でフィルタ ===
def filter_sales(df: pd.DataFrame, category: str, dates: list[str]) -> pd.DataFrame:
    df = df[df["category_large"].astype(str) == str(category)]
    dates = [pd.to_datetime(d).date() for d in dates]
    df = df[df["date"].isin(dates)]
    return df

# === 店舗×日付TopN抽出 ===
def build_topn(df: pd.DataFrame, top_n=30) -> dict:
    result = {}
    for (store, date), g in df.groupby(["store_id", "date"]):
        g = g.sort_values("amount", ascending=False).head(top_n)
        total = g["amount"].sum()
        g["構成比"] = g["amount"] / total if total else 0
        result.setdefault(store, {})[date] = g
    return result

# === Excel書き出し ===
def write_excel(
    template_path,
    out_path,
    topn_dict,
    store_names,
    category,
    dates,
    event_name="イベント名",
    df_sales_all=None,  # ←追加
):

    wb = load_workbook(template_path)
    ws_tpl = wb["TEMPLATE"]

    cat_name = CATEGORY_MAP.get(str(category), str(category))
    total_days = len(dates)
    num_pages = math.ceil(total_days / 4)
    block_offsets = [0, 8, 16, 24]  # 各ブロックの列オフセット

    for store in sorted(topn_dict.keys(), key=lambda x: int(x)):
        day_map = topn_dict[store]
        short_name = store_names.get(store, "")
        for page in range(num_pages):
            ws = wb.copy_worksheet(ws_tpl)
            copy_conditional_formatting(ws, ws_tpl)
            ws.title = f"{store}({page+1})"
            # 該当ページの4日分
            page_dates = dates[page*4 : (page+1)*4]

            # ==== 各ブロックごと ====
            for block_idx, d in enumerate(page_dates):
                if block_idx >= 4:
                    break
                col_offset = block_offsets[block_idx]
                df_day = day_map.get(pd.to_datetime(d).date())
                if df_day is None or df_day.empty:
                    continue

                # === タイトル A1 ===
                title_text = f"{event_name}　{d}　{cat_name}単品データ"
                ws["A1"].value = title_text

                # === ブロックヘッダ ===
                year = str(pd.to_datetime(d).year)[2:]
                month_day = pd.to_datetime(d).strftime("%m/%d")
                ws.cell(row=2, column=1+col_offset, value=year)
                ws.cell(row=2, column=2+col_offset, value=month_day)
                ws.cell(row=2, column=3+col_offset, value=short_name)
                ws.cell(row=2, column=4+col_offset, value=f"{cat_name}単品")

                # === 見出し ===
                headers = ["順位", "商品名", "売上金額", "売上数量", "値引金額", "値引率"]
                for i, h in enumerate(headers):
                    ws.cell(row=3, column=1+col_offset+i, value=h)

                # === TopNデータ (最大35行) ===
                # 事前に金額で降順
                df_day = df_day.sort_values("amount", ascending=False)

                for rank, row in enumerate(df_day.itertuples(index=False), start=1):
                    if rank > 35:
                        break
                    r = 3 + rank

                    # 先に数値を取り出してから値引率を計算
                    amt  = float(getattr(row, "amount", 0) or 0.0)
                    qty  = float(getattr(row, "qty", 0) or 0.0)
                    disc = float(getattr(row, "discount", 0) or 0.0)
                    rate_val = (disc / amt) if amt else 0.0

                    ws.cell(r, 1+col_offset, rank)          # 順位
                    ws.cell(r, 2+col_offset, getattr(row, "name", ""))  # 商品名
                    ws.cell(r, 3+col_offset, amt)           # 売上金額
                    ws.cell(r, 4+col_offset, qty)           # 売上数量
                    ws.cell(r, 5+col_offset, disc)          # 値引金額
                    cell = ws.cell(r, 6+col_offset, rate_val)  # 値引率（値を書き込む）
                    cell.number_format = "0.00%"


                # === フッタ（合計行） ===
                row_base = 39
                ws.cell(row=row_base+1, column=1+col_offset, value="惣菜売上金額")
                ws.cell(row=row_base+2, column=1+col_offset, value=f"{cat_name}売上金額")
                ws.cell(row=row_base+3, column=1+col_offset, value=f"{cat_name}構成比")

                # ① 惣菜売上金額＝全カテゴリ合計（全データから抽出）
                if df_sales_all is not None:
                    total_store_amount = (
                        df_sales_all[
                            (df_sales_all["store_id"] == store)
                            & (df_sales_all["date"] == pd.to_datetime(d).date())
                        ]["amount"]
                        .sum()
                    )
                else:
                    total_store_amount = df_day["amount"].sum()

                # ② 大分類売上金額（TopN以外も含む全アイテムの合計）
                if df_sales_all is not None:
                    total_cat_amount = (
                        df_sales_all[
                            (df_sales_all["store_id"] == store)
                            & (df_sales_all["date"] == pd.to_datetime(d).date())
                            & (df_sales_all["category_large"].astype(str) == str(category))
                        ]["amount"]
                        .sum()
                    )
                else:
                    total_cat_amount = df_day["amount"].sum()

                # 構成比
                ratio = total_cat_amount / total_store_amount if total_store_amount else 0

                ws.cell(row=row_base+1, column=3+col_offset, value=total_store_amount)
                ws.cell(row=row_base+2, column=3+col_offset, value=total_cat_amount)
                ws.cell(row=row_base+3, column=3+col_offset, value=ratio)
                ws.cell(row=row_base+3, column=3+col_offset).number_format = "0.00%"


    del wb["TEMPLATE"]
    wb.save(out_path)
    print(f"[ok] saved → {out_path}")

if __name__ == "__main__":
    print("[debug] 開始")

    root = Path("data/material")
    df = load_sales(root)
    print(f"[debug] sales rows={len(df)}")

    store_names = load_store_master(root / "master" / "store_master.xlsx")
    print(f"[debug] stores={len(store_names)}")

    category = "4"  # 冷総菜
    dates = ["2025-01-03", "2025-01-10", "2025-01-17", "2025-01-24"]

    df_f = filter_sales(df, category, dates)
    print(f"[debug] filtered rows={len(df_f)}")

    topn = build_topn(df_f, top_n=35)
    print(f"[debug] topn stores={len(topn)}")

    write_excel(
        template_path="data/template/配布フォーマット.xlsx",
        out_path="data/output/topN_冷総菜_202501.xlsx",
        topn_dict=topn,
        store_names=store_names,
        category=category,
        dates=dates,
        event_name="秋の感謝セール",
        df_sales_all=df
    )
