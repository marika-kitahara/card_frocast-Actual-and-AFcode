import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="実績データ集計", layout="wide")

AF_MASTER_PATH = "AFマスター.xlsx"
AFF_ONLY_PATH = "AFF_AFコード.xlsx"
TARGET_APPLY_PATH = "目標申込件数マスター.xlsx"
TARGET_ISSUE_PATH = "目標発行件数マスター.xlsx"

DM_CODES = {
    "MCS056",
    "APD023",
    "GEZ039",
    "STD002"
}

AREA_ORDER = [
    "Affiliate",
    "Listing",
    "Display",
    "DM"
]

AREA_NAME_MAP = {
    "affiliate": "Affiliate",
    "listing": "Listing",
    "display": "Display",
    "dm": "DM"
}

# ======================
# 共通関数
# ======================
def normalize(val):
    if pd.isna(val):
        return ""
    return str(val).replace(" ", "").replace("　", "").strip()

def convert_date(val):
    try:
        s = str(int(val))
        return pd.Timestamp(f"{s[:4]}/{int(s[4:6])}/{int(s[6:8])}")
    except:
        return pd.NaT

# ======================
# マスタ読込
# ======================
def read_af_master(path):
    df = pd.read_excel(path, header=None)

    header_row = (
        df.apply(lambda r: r.astype(str).str.contains("AFコード")).any(axis=1).idxmax()
    )

    df.columns = df.iloc[header_row]
    df = df.iloc[header_row + 1:].reset_index(drop=True)

    df.columns = df.columns.map(normalize)
    df["AFコード"] = df["AFコード"].apply(normalize)
    df["割り振り"] = df["割り振り"].apply(normalize)

    return df[["AFコード", "割り振り", "領域"]]

def read_aff_master(path):
    df = pd.read_excel(path)

    if isinstance(df, pd.Series):
        df = df.to_frame()

    df.columns = ["AFコード", "領域"]
    df["割り振り"] = df["AFコード"]

    for col in df.columns:
        df[col] = df[col].apply(normalize)

    return df

def read_target(path):
    df = pd.read_excel(path, header=4)
    df.columns = df.columns.map(normalize)
    df.rename(columns={df.columns[1]: "日付"}, inplace=True)
    df["日付"] = pd.to_datetime(df["日付"])
    return df

def get_target(df, date, assign):
    row = df[df["日付"] == date]
    if row.empty or assign not in df.columns:
        return 0
    v = row.iloc[0][assign]
    return 0 if pd.isna(v) else v

# ======================
# ローデータ処理
# ======================
def process_raw(df_raw, af_master, start, end, kind):
    af_map = (
        af_master
        .set_index("AFコード")[["割り振り", "領域"]]
        .to_dict("index")
    )

    rows = []

    for _, r in df_raw.iterrows():
        if pd.isna(r["日付"]) or not (start <= r["日付"] <= end):
            continue

        for col in df_raw.columns[1:]:
            if pd.isna(r[col]) or r[col] == 0:
                continue

            code = normalize(col)
            info = af_map.get(code)

            # DMコードはマスター未登録でもDMとして集計
            if code in DM_CODES:
                rows.append(
                    [kind, r["日付"], code, "DM", r[col]]
                )
                continue

            if info:
                rows.append(
                    [kind, r["日付"], info["割り振り"], info["領域"], r[col]]
                )

    df = pd.DataFrame(
        rows,
        columns=["種別", "日付", "割り振り", "領域", "実績"]
    )

    return (
        df.groupby(["種別", "日付", "割り振り", "領域"], as_index=False)
        .sum()
    )

# ===============================
# 割り振り別ブロック（修正版）
# ===============================
def create_blocks(df, target):
    act = df.pivot_table(
        index="日付", columns="割り振り", values="実績",
        aggfunc="sum", fill_value=0
    )

    # ✅ ここ超重要
    act = act.astype(float)
    tar = act.copy()

    for d in act.index:
        for c in act.columns:
            tar.loc[d, c] = float(get_target(target, d, c))

    for t in (act, tar):
        t["total"] = t.sum(axis=1)
        t.index = t.index.strftime("%Y/%m/%d")

    gap = act - tar
    ratio = act.divide(tar.replace(0, pd.NA)).fillna(0)

    return act, tar, gap, ratio

# ======================
# Excel 出力
# ======================
def to_excel(area_a, area_i, blocks_a, blocks_i, raw):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:

        # ---- 割り振り別 ----
        for label, blocks in [("申込", blocks_a), ("発行", blocks_i)]:
            for name, df in zip(["実績", "目標", "GAP", "比率"], blocks):
                df.to_excel(writer, sheet_name=f"{label}_{name}")

        raw.to_excel(writer, sheet_name="ローデータ", index=False)

    return output.getvalue()

# ======================
# UI
# ======================
st.title("📊 実績データ集計")

apply = st.file_uploader("📤 申込データ", type="xlsx")
issue = st.file_uploader("📤 発行データ", type="xlsx")

if not apply or not issue:
    st.stop()

dfa = pd.read_excel(apply)
dfi = pd.read_excel(issue)

dfa.rename(columns={dfa.columns[0]: "日付"}, inplace=True)
dfi.rename(columns={dfi.columns[0]: "日付"}, inplace=True)

dfa["日付"] = dfa["日付"].apply(convert_date)
dfi["日付"] = dfi["日付"].apply(convert_date)

date_range = st.date_input(
    "📅 期間選択",
    [dfa["日付"].min(), dfa["日付"].max()]
)

if len(date_range) != 2:
    st.stop()

start, end = map(pd.to_datetime, date_range)

af = pd.concat(
    [read_af_master(AF_MASTER_PATH), read_aff_master(AFF_ONLY_PATH)],
    ignore_index=True
)

ta = read_target(TARGET_APPLY_PATH)
ti = read_target(TARGET_ISSUE_PATH)

ra = process_raw(dfa, af, start, end, "申込")
ri = process_raw(dfi, af, start, end, "発行")

# ---- サマリ ----
st.subheader("📌 サマリ（領域別）")

def ensure_area_rows(area_df):
    df = area_df.copy()

    # 領域名の表記ゆれを統一する
    df.index = [
        AREA_NAME_MAP.get(normalize(idx).lower(), normalize(idx))
        for idx in df.index
    ]

    # 表記ゆれで同じ領域が複数できた場合は合算する
    df = df.groupby(level=0).sum()

    # DMなど、実績がない領域も0行として追加する
    for area in AREA_ORDER:
        if area not in df.index:
            df.loc[area] = 0

    # 表示順を固定する
    extra_areas = [idx for idx in df.index if idx not in AREA_ORDER]
    return df.reindex(AREA_ORDER + extra_areas, fill_value=0)


def make_summary_df(area_df):
    df = area_df.copy()
    df.insert(0, "total", df.sum(axis=1))
    df.loc["total"] = df.sum()
    df.columns = ["total"] + [
        pd.to_datetime(c).strftime("%Y/%m/%d") for c in df.columns[1:]
    ]
    return df

st.markdown("### 申込")
area_apply = ra.pivot_table(
    index="領域",
    columns="日付",
    values="実績",
    aggfunc="sum",
    fill_value=0
)
area_apply = ensure_area_rows(area_apply)
st.dataframe(make_summary_df(area_apply), use_container_width=True)

st.markdown("### 発行")
area_issue = ri.pivot_table(
    index="領域",
    columns="日付",
    values="実績",
    aggfunc="sum",
    fill_value=0
)
area_issue = ensure_area_rows(area_issue)
st.dataframe(make_summary_df(area_issue), use_container_width=True)

# ---- ローデータ ----
raw = pd.concat([ra, ri], ignore_index=True)
raw["目標"] = raw.apply(
    lambda r: get_target(
        ta if r["種別"] == "申込" else ti,
        r["日付"],
        r["割り振り"]
    ),
    axis=1
)
raw["日付"] = raw["日付"].dt.strftime("%Y/%m/%d")

excel = to_excel(
    area_apply,
    area_issue,
    create_blocks(ra, ta),
    create_blocks(ri, ti),
    raw
)

st.download_button(
    "📥 集計結果をダウンロード（Excel）",
    excel,
    f"実績集計結果_{start:%Y%m%d}_{end:%Y%m%d}.xlsx",
    use_container_width=True
)
