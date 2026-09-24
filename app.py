import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="実績データ集計", layout="wide")

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

# ======================
# ローデータ処理
# ======================
def process_raw(df_raw, af_master, start, end, kind):
    # AFマスタとAFFマスタを結合した際に同じAFコードが複数存在しても
    # to_dict("index") でエラーにならないよう、AFコードを一意化する。
    # 先に読み込んでいるAFマスタ側を優先する。
    af_unique = (
        af_master
        .copy()
        .loc[lambda x: x["AFコード"] != ""]
        .drop_duplicates(subset=["AFコード"], keep="first")
    )

    af_map = (
        af_unique
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
# 割り振り別実績ブロック
# ===============================
def create_actual_block(df):
    act = df.pivot_table(
        index="日付", columns="割り振り", values="実績",
        aggfunc="sum", fill_value=0
    ).astype(float)

    act["total"] = act.sum(axis=1)
    act.index = act.index.strftime("%Y/%m/%d")
    return act

# ======================
# Excel 出力
# ======================
def to_excel(
    actual_a,
    actual_i,
    raw,
    summary_apply,
    summary_issue
):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:

        # ======================
        # サマリシート
        # ======================
        sheet_name = "サマリ"

        row = 0

        pd.DataFrame([["申込"]]).to_excel(
            writer,
            sheet_name=sheet_name,
            startrow=row,
            index=False,
            header=False
        )

        row += 1

        summary_apply.to_excel(
            writer,
            sheet_name=sheet_name,
            startrow=row
        )

        row += len(summary_apply) + 4

        pd.DataFrame([["発行"]]).to_excel(
            writer,
            sheet_name=sheet_name,
            startrow=row,
            index=False,
            header=False
        )

        row += 1

        summary_issue.to_excel(
            writer,
            sheet_name=sheet_name,
            startrow=row
        )

        # ======================
        # 割り振り別実績
        # ======================
        actual_a.to_excel(
            writer,
            sheet_name="申込_実績"
        )

        actual_i.to_excel(
            writer,
            sheet_name="発行_実績"
        )

        # ======================
        # ローデータ
        # ======================
        raw.to_excel(
            writer,
            sheet_name="ローデータ",
            index=False
        )

    return output.getvalue()

# ======================
# UI
# ======================
st.title("📊 実績データ集計")

st.subheader("マスタファイル")

af_master_file = st.file_uploader(
    "📤 AFマスタ",
    type=["xlsx"],
    key="af_master"
)

st.markdown(
    "[📂 AFマスタはこちら]"
    "(https://rak.box.com/s/jmks0tjanuskp957lq4wmepun5cpm5xd)"
)

aff_master_file = st.file_uploader(
    "📤 AFFマスタ",
    type=["xlsx"],
    key="aff_master"
)

st.markdown(
    "[📂 AFFマスタはこちら]"
    "(https://rak.box.com/s/rtkp5rshiwqsa69pkezl13881b552oe0)"
)

st.subheader("実績データ")

apply = st.file_uploader(
    "📤 申込データ",
    type=["xlsx"],
    key="apply"
)

issue = st.file_uploader(
    "📤 発行データ",
    type=["xlsx"],
    key="issue"
)
required_files = {
    "AFマスタ": af_master_file,
    "AFFマスタ": aff_master_file,
    "申込データ": apply,
    "発行データ": issue,
}

missing_files = [
    name
    for name, uploaded_file in required_files.items()
    if uploaded_file is None
]

if missing_files:
    st.info(
        "次の必須ファイルをアップロードしてください："
        + "、".join(missing_files)
    )
    st.stop()
    
dfa = pd.read_excel(apply)
dfi = pd.read_excel(issue)

dfa.rename(columns={dfa.columns[0]: "日付"}, inplace=True)
dfi.rename(columns={dfi.columns[0]: "日付"}, inplace=True)

dfa["日付"] = dfa["日付"].apply(convert_date)
dfi["日付"] = dfi["日付"].apply(convert_date)

# 申込・発行データのうち、存在する最新日を基準に
# デフォルト期間を「最新日の月の1日 ～ 最新日」に固定する。
all_dates = pd.concat([dfa["日付"], dfi["日付"]], ignore_index=True).dropna()

if all_dates.empty:
    st.error("申込データ・発行データに有効な日付がありません。")
    st.stop()

latest_date = all_dates.max()
default_start = latest_date.replace(day=1)

date_range = st.date_input(
    "📅 期間選択",
    value=(default_start.date(), latest_date.date()),
    min_value=all_dates.min().date(),
    max_value=latest_date.date(),
)

if len(date_range) != 2:
    st.stop()

start, end = map(pd.to_datetime, date_range)

try:
    af = pd.concat(
        [
            read_af_master(af_master_file),
            read_aff_master(aff_master_file),
        ],
        ignore_index=True
    )
except Exception as e:
    st.error(f"AF・AFFマスタの読み込みに失敗しました：{e}")
    st.stop()


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

summary_apply = make_summary_df(area_apply)

st.dataframe(
    summary_apply,
    use_container_width=True
)
st.markdown("### 発行")
area_issue = ri.pivot_table(
    index="領域",
    columns="日付",
    values="実績",
    aggfunc="sum",
    fill_value=0
)
area_issue = ensure_area_rows(area_issue)

summary_issue = make_summary_df(area_issue)

st.dataframe(
    summary_issue,
    use_container_width=True
)
# ---- ローデータ ----
raw = pd.concat([ra, ri], ignore_index=True)
raw["日付"] = raw["日付"].dt.strftime("%Y/%m/%d")

excel = to_excel(
    create_actual_block(ra),
    create_actual_block(ri),
    raw,
    summary_apply,
    summary_issue
)

st.download_button(
    "📥 集計結果をダウンロード（Excel）",
    excel,
    f"実績集計結果_{start:%Y%m%d}_{end:%Y%m%d}.xlsx",
    use_container_width=True
)
