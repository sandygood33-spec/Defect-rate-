"""
零件耗用一覽 / 故障率查詢工具
================================
使用方式:
    streamlit run app.py

部署:
    放到 GitHub（免費 repo 即可），到 https://share.streamlit.io 用該 repo 建立
    一個 Streamlit Community Cloud 應用程式，即可得到一個瀏覽器網址給同事使用。

密碼保護:
    在 Streamlit Cloud 的 App settings -> Secrets 裡加入：
        APP_PASSWORD = "你自訂的密碼"
    本機測試時，可在 streamlit_app/.streamlit/secrets.toml 裡加同一行。

★★★ 2026/08 更新（第二輪）：已用真實檔案核對過欄位與邏輯，重要變更如下：

    1. 「故障部位」：耗用資料裡雖然有現成的「故障部位」欄位（例如
       "C02 控制PCB"），但依照使用者指示，**忽略這個原始欄位**，改為用
       「零件編號」對照「零件分類一覽表」的「分類」欄位（壓縮機/機板/
       風扇馬達…31種）來決定故障部位。找不到對應分類的零件編號，會標
       為「未分類（新零件）」，並在「新零件料號提醒」裡一併列出。
    2. 耗用資料裡有些列「零件編號」是空的、「數量」是 0（像是同一張
       維修單底下的參考列），這些不算真的耗用，程式會先濾掉。
    3. 「維修單號+零件編號」去重複的邏輯是錯的——同一張維修單本來就
       可能合法地用到同一個零件兩次，已拿掉，改成單純合併兩個檔案
       （因為兩份月資料的故障日期本來就不重疊）。
    4. 出貨資料的 Excel 有多個工作表，且正確工作表的「分頁名稱」每月
       會變動（例如 9 月會變成 20210101-20260831 這種格式），所以不能
       用固定分頁名稱抓取。程式改成自動偵測「有機型欄 + 販售量欄」的
       那個分頁，不受名稱變動影響。
"""

import io
import re
from datetime import date, datetime

import pandas as pd
import streamlit as st

st.set_page_config(page_title="零件耗用 / 故障率查詢", layout="wide")

# ---------------------------------------------------------------------------
# 假設的欄位名稱（依 PPT 截圖推測，上線前務必用真實檔案核對）
# ---------------------------------------------------------------------------
ASSUMED_USAGE_COLUMNS = {
    "repair_no": "維修單號",       # 用來判讀維修分公司別（第5碼）
    "model": "維修機型",           # 對應機型分類表的「機型」
    "install_date": "裝機日",
    "fault_date": "故障日",        # 所有邏輯以此欄位為基準
    "fault_category": "故障類別",
    "fault_part_raw": "故障部位",  # 耗用資料自帶欄位，依指示不使用，只保留供除錯參考
    "part_no": "零件編號",         # 對應零件分類一覽表的「零件編號」，用來決定故障部位＋新零件提醒
    "part_name": "品名",
    "qty": "數量",
}

# 出貨資料：機型 + 每年一欄（例如 "2021販售量", ..., "2026販售量(-0731)"）
SHIPMENT_MODEL_COL = "機型"
SHIPMENT_YEAR_COL_PATTERN = re.compile(r"^(20\d{2})販售量")

BRANCH_CODE_MAP = {
    "0": "台北公司", "2": "台北公司", "3": "桃園分公司", "4": "新竹分公司",
    "5": "中部分公司", "6": "嘉義分公司", "7": "台南分公司", "8": "高雄分公司",
    "B": "花蓮分公司", "C": "宜蘭分公司", "E": "屏東分公司", "F": "基隆分公司",
}

USAGE_TABLE_COLS = [
    "類別", "室內/外機", "機型", "故障部位", "零件料號", "維修分公司", "當期耗用量",
]
RATE_TABLE_COLS = [
    "類別", "室內/外機", "機型", "故障部位", "零件料號", "維修分公司",
    "該月累積故障率", "前月累積故障率", "去年同期累積故障率",
]


# ---------------------------------------------------------------------------
# 密碼保護
# ---------------------------------------------------------------------------
def check_password() -> bool:
    def _password_entered():
        correct = st.secrets.get("APP_PASSWORD", None)
        if correct is None:
            st.session_state["password_ok"] = True  # 沒設密碼時，本機測試放行
            return
        st.session_state["password_ok"] = (
            st.session_state.get("password_input", "") == correct
        )

    if st.session_state.get("password_ok"):
        return True

    st.title("零件耗用 / 故障率查詢")
    st.text_input("請輸入密碼", type="password", key="password_input",
                   on_change=_password_entered)
    if st.session_state.get("password_ok") is False:
        st.error("密碼錯誤，請再試一次")
    return False


# ---------------------------------------------------------------------------
# 資料讀取與清理
# ---------------------------------------------------------------------------
def read_excel(uploaded_file) -> pd.DataFrame:
    return pd.read_excel(uploaded_file)


def validate_columns(df: pd.DataFrame, required: list, label: str) -> list:
    missing = [c for c in required if c not in df.columns]
    if missing:
        st.error(
            f"「{label}」缺少欄位：{missing}。目前偵測到的欄位是：{list(df.columns)}。"
            f"請確認檔案格式，或告訴我正確欄位名稱讓我調整程式。"
        )
    return missing


def branch_from_repair_no(repair_no: str) -> str:
    if not isinstance(repair_no, str) or len(repair_no) < 5:
        return "未知"
    code = repair_no[4]  # 第5碼，index 4
    return BRANCH_CODE_MAP.get(code.upper(), "未知")


def extract_year_month(d) -> str:
    """回傳 'YYYY-MM'；輸入可能已經是 datetime，也可能是文字。"""
    if pd.isna(d):
        return None
    if isinstance(d, (pd.Timestamp, datetime, date)):
        return f"{d.year:04d}-{d.month:02d}"
    s = str(d)
    m = re.match(r"(\d{4})[-/](\d{1,2})", s)
    if m:
        return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}"
    return None


@st.cache_data(show_spinner=False)
def build_usage_df(usage_recent_bytes, usage_history_bytes,
                    model_map_bytes, part_map_bytes):
    """usage_recent / usage_history：兩份零件耗用資料，故障日範圍不重疊，
    直接合併即可，不需要去重（同一張維修單合法地用到同一個零件兩次的情況
    是存在的，用維修單號+零件編號去重會誤刪真實資料）。"""
    df_a = pd.read_excel(io.BytesIO(usage_recent_bytes))
    df_b = pd.read_excel(io.BytesIO(usage_history_bytes))
    model_map = pd.read_excel(io.BytesIO(model_map_bytes))
    part_map = pd.read_excel(io.BytesIO(part_map_bytes))

    cols = ASSUMED_USAGE_COLUMNS
    for df, label in [(df_a, "零件耗用資料（近期）"), (df_b, "零件耗用資料（歷史）")]:
        validate_columns(df, list(cols.values()), label)
    validate_columns(model_map, ["類別", "機型", "內外機"], "機型分類表")
    validate_columns(part_map, ["分類", "零件編號", "品名"], "零件分類一覽表")

    usage = pd.concat([df_b, df_a], ignore_index=True)

    # 新零件料號提醒：耗用資料裡有、但零件分類一覽表沒有的零件編號
    # （在濾掉空白零件編號之前先算，這樣提醒名單才完整）
    known_parts = set(part_map["零件編號"].astype(str))
    used_parts_all = set(usage[cols["part_no"]].dropna().astype(str))
    unknown_parts = sorted(used_parts_all - known_parts)

    # 排除零件編號空白、數量為0的列（不是真的耗用，是同一張維修單的參考列）
    usage = usage[usage[cols["part_no"]].notna() & (usage[cols["qty"]] > 0)].copy()

    usage["年月"] = usage[cols["fault_date"]].apply(extract_year_month)
    usage["維修分公司"] = usage[cols["repair_no"]].apply(branch_from_repair_no)

    # 關聯機型分類表 -> 類別 / 內外機
    model_lookup = model_map.set_index("機型")[["類別", "內外機"]]
    usage = usage.join(model_lookup, on=cols["model"])
    usage = usage.rename(columns={cols["model"]: "機型"})

    # 故障部位：依指示，用零件編號對照「零件分類一覽表」的「分類」欄位決定，
    # 不使用耗用資料自帶的「故障部位」欄位。對照不到的零件編號（新零件）
    # 標成「未分類（新零件）」，讓使用者在畫面上也看得到，而不是悄悄消失。
    part_lookup = part_map.set_index("零件編號")["分類"]
    usage["故障部位"] = usage[cols["part_no"]].astype(str).map(part_lookup)
    usage["故障部位"] = usage["故障部位"].fillna("未分類（新零件）")
    usage = usage.rename(columns={cols["part_no"]: "零件料號"})
    usage["零件料號"] = usage["零件料號"].astype(str)

    return usage, unknown_parts, model_map, part_map


@st.cache_data(show_spinner=False)
def build_shipment_df(shipment_bytes):
    xl = pd.ExcelFile(io.BytesIO(shipment_bytes))
    chosen = None
    for name in xl.sheet_names:
        tmp = xl.parse(name, nrows=3)
        if SHIPMENT_MODEL_COL in tmp.columns and any(
            SHIPMENT_YEAR_COL_PATTERN.match(str(c)) for c in tmp.columns
        ):
            chosen = name
            break
    if chosen is None:
        chosen = xl.sheet_names[0]
    df = xl.parse(chosen)
    validate_columns(df, [SHIPMENT_MODEL_COL], "出貨資料")
    year_cols = [c for c in df.columns if SHIPMENT_YEAR_COL_PATTERN.match(str(c))]
    if not year_cols:
        st.warning("在出貨資料裡找不到「20XX販售量」這種格式的欄位，累積出貨量會算不出來，"
                   "請確認出貨資料的欄位名稱。")
    return df, year_cols


# ---------------------------------------------------------------------------
# 業務邏輯：零件耗用一覽
# ---------------------------------------------------------------------------
def shift_month(ym: str, months: int) -> str:
    y, m = map(int, ym.split("-"))
    idx = y * 12 + (m - 1) + months
    return f"{idx // 12:04d}-{idx % 12 + 1:02d}"


def usage_table(usage_df, start_ym, end_ym, f_cat, f_io, f_model, f_part, f_partno, f_branch):
    """查找年月(起始~結束) 整段期間內的耗用量加總。
    2026/08 更新：不再算「前月/去年同期」單月數字，改成直接加總使用者選的
    查找區間本身（例如選 2021-01~2026-07，就是這段期間的總耗用量）。"""
    scope = usage_df[(usage_df["年月"] >= start_ym) & (usage_df["年月"] <= end_ym)]
    if f_cat != "全部": scope = scope[scope["類別"] == f_cat]
    if f_io != "全部": scope = scope[scope["內外機"] == f_io]
    if f_model != "全部": scope = scope[scope["機型"] == f_model]
    if f_part != "全部": scope = scope[scope["故障部位"] == f_part]
    if f_partno != "全部": scope = scope[scope["零件料號"] == f_partno]
    if f_branch != "全部": scope = scope[scope["維修分公司"] == f_branch]

    group_cols = ["類別", "內外機", "機型", "故障部位", "零件料號", "維修分公司"]
    result = (
        scope.groupby(group_cols, as_index=False)[ASSUMED_USAGE_COLUMNS["qty"]]
        .sum()
        .rename(columns={"內外機": "室內/外機", ASSUMED_USAGE_COLUMNS["qty"]: "當期耗用量"})
    )
    return result[USAGE_TABLE_COLS]


# ---------------------------------------------------------------------------
# 業務邏輯：故障率
# ---------------------------------------------------------------------------
def cumulative_shipment(shipment_df, year_cols, model, end_ym):
    """2021/01/01 ~ end_ym 月底 的累積出貨量（受限於出貨資料是年度粒度）"""
    end_year = int(end_ym.split("-")[0])
    row = shipment_df[shipment_df[SHIPMENT_MODEL_COL] == model]
    if row.empty:
        return 0
    total = 0
    for c in year_cols:
        y = int(SHIPMENT_YEAR_COL_PATTERN.match(c).group(1))
        if y <= end_year:
            total += row[c].sum()
    return total


def _filter_scope(usage_df, f_cat, f_io, f_model, f_part, f_partno, f_branch):
    scope = usage_df
    if f_cat != "全部": scope = scope[scope["類別"] == f_cat]
    if f_io != "全部": scope = scope[scope["內外機"] == f_io]
    if f_model != "全部": scope = scope[scope["機型"] == f_model]
    if f_part != "全部": scope = scope[scope["故障部位"] == f_part]
    if f_partno != "全部": scope = scope[scope["零件料號"] == f_partno]
    if f_branch != "全部": scope = scope[scope["維修分公司"] == f_branch]
    return scope


def rate_table(usage_df, shipment_df, year_cols, end_ym,
                f_cat, f_io, f_model, f_part, f_partno, f_branch):
    """2026/08 效能重寫：原本用 apply() 對每個組合逐一重新掃描整個 usage_df 三次
    （該月/前月/去年同期各一次），資料量一大會非常慢。改成：先把使用者選的
    條件套用一次（scope），再用 groupby 一次算出每個組合在三個月份的累積量，
    整體只需要掃描 scope 三次，不是「組合數 × 3」次。"""
    scope = _filter_scope(usage_df, f_cat, f_io, f_model, f_part, f_partno, f_branch)
    group_cols = ["類別", "內外機", "機型", "故障部位", "零件料號", "維修分公司"]
    combos = scope[group_cols].drop_duplicates()

    empty_cols = RATE_TABLE_COLS
    if combos.empty:
        return pd.DataFrame(columns=empty_cols)

    prev_ym = shift_month(end_ym, -1)
    last_year_ym = shift_month(end_ym, -12)

    def cumulative_by_combo(target_ym):
        sub = scope[scope["年月"] <= target_ym]
        if sub.empty:
            return pd.Series(dtype=float)
        return sub.groupby(group_cols)[ASSUMED_USAGE_COLUMNS["qty"]].sum()

    cum_end = cumulative_by_combo(end_ym)
    cum_prev = cumulative_by_combo(prev_ym)
    cum_lastyear = cumulative_by_combo(last_year_ym)

    combos = combos.set_index(group_cols)
    combos["_end_qty"] = cum_end.reindex(combos.index).fillna(0)
    combos["_prev_qty"] = cum_prev.reindex(combos.index).fillna(0)
    combos["_lastyear_qty"] = cum_lastyear.reindex(combos.index).fillna(0)
    combos = combos.reset_index()

    unique_models = combos["機型"].unique().tolist()
    ship_end = {m: cumulative_shipment(shipment_df, year_cols, m, end_ym) for m in unique_models}
    ship_prev = {m: cumulative_shipment(shipment_df, year_cols, m, prev_ym) for m in unique_models}
    ship_lastyear = {m: cumulative_shipment(shipment_df, year_cols, m, last_year_ym) for m in unique_models}

    def safe_rate(qty, ship):
        return (qty / ship) if ship else None

    combos["該月累積故障率"] = combos.apply(lambda r: safe_rate(r["_end_qty"], ship_end[r["機型"]]), axis=1)
    combos["前月累積故障率"] = combos.apply(lambda r: safe_rate(r["_prev_qty"], ship_prev[r["機型"]]), axis=1)
    combos["去年同期累積故障率"] = combos.apply(lambda r: safe_rate(r["_lastyear_qty"], ship_lastyear[r["機型"]]), axis=1)
    combos = combos.rename(columns={"內外機": "室內/外機"})

    for c in ["該月累積故障率", "前月累積故障率", "去年同期累積故障率"]:
        combos[c] = combos[c].apply(lambda v: f"{v:.2%}" if pd.notna(v) else "-")
    return combos[RATE_TABLE_COLS]


# ---------------------------------------------------------------------------
# 畫面
# ---------------------------------------------------------------------------
def cascading_select(label, options, key, disabled=False):
    opts = ["全部"] + [o for o in options if o != "全部"]
    return st.selectbox(label, opts, key=key, disabled=disabled)


def safe_selectbox(label, options, key, disabled=False, **kwargs):
    """2026/08 修 bug：上游條件（類別/內外機）改變時，下游選單（機型/零件料號）
    的選項列表也會跟著變，但 Streamlit 的 selectbox 若沿用同一個 key，殘留的
    舊選擇值可能不在新的選項列表裡（例如原本選「室內機」+「RHF25RVLT」，
    後來把室內外機改成別的，機型卻沒跟著清空），造成篩選條件互相矛盾、
    查不到任何資料。這裡在畫出選單前先檢查目前的殘留值還在不在新選項裡，
    不在的話強制重置成「全部」，避免出現不可能存在的組合。"""
    if key in st.session_state and st.session_state[key] not in options:
        st.session_state[key] = "全部" if "全部" in options else (options[0] if options else "")
    return st.selectbox(label, options, key=key, disabled=disabled, **kwargs)


def main():
    if not check_password():
        return

    st.title("零件耗用 / 故障率查詢")

    with st.expander("📤 上傳資料（每次都請上傳最新檔案）", expanded=True):
        c1, c2 = st.columns(2)
        with c1:
            f_usage_year = st.file_uploader("① 最新一期零件耗用資料（例如當月）", type="xlsx")
            f_usage_2021 = st.file_uploader("② 歷史零件耗用資料（2021~上一期為止）", type="xlsx")
            f_shipment = st.file_uploader("③ 2021~至今出貨資料", type="xlsx")
        with c2:
            f_model_map = st.file_uploader("④ 機型分類表（機型分類_已填寫）", type="xlsx")
            f_part_map = st.file_uploader("⑤ 零件分類一覽表", type="xlsx")

    if not all([f_usage_year, f_usage_2021, f_shipment, f_model_map, f_part_map]):
        st.info("請上傳以上 5 份檔案後才能開始查詢。")
        return

    usage_df, unknown_parts, model_map, part_map = build_usage_df(
        f_usage_year.getvalue(), f_usage_2021.getvalue(),
        f_model_map.getvalue(), f_part_map.getvalue(),
    )
    shipment_df, year_cols = build_shipment_df(f_shipment.getvalue())

    if unknown_parts:
        st.warning(
            f"⚠️ 有 {len(unknown_parts)} 個零件料號沒有登錄在「零件分類一覽表」裡，"
            f"請確認是否為新零件：{', '.join(unknown_parts[:30])}"
            + ("...（僅顯示前30個）" if len(unknown_parts) > 30 else "")
        )

    all_cats = sorted(model_map["類別"].dropna().unique().tolist())
    all_ios = sorted(model_map["內外機"].dropna().unique().tolist())
    all_parts = sorted(part_map["分類"].dropna().unique().tolist())
    all_branches = sorted(BRANCH_CODE_MAP.values())

    tab_usage, tab_rate = st.tabs(["零件耗用一覽", "故障率"])

    # ---------------- 零件耗用一覽 ----------------
    with tab_usage:
        st.subheader("查詢條件")
        r1 = st.columns(4)
        start_ym = r1[0].text_input("查找年月（起始，YYYY-MM）", value="2021-01", key="u_start")
        end_ym = r1[1].text_input("查找年月（結束，YYYY-MM）", value="2026-07", key="u_end")
        f_cat = r1[2].selectbox("類別", ["全部"] + all_cats, key="u_cat")
        f_io = r1[3].selectbox("室內/外機", ["全部"] + all_ios, key="u_io")

        model_scope = model_map
        if f_cat != "全部": model_scope = model_scope[model_scope["類別"] == f_cat]
        if f_io != "全部": model_scope = model_scope[model_scope["內外機"] == f_io]
        avail_models = sorted(model_scope["機型"].unique().tolist())

        r2 = st.columns(4)
        f_model = safe_selectbox("機型", ["全部"] + avail_models, key="u_model")
        f_part = r2[1].selectbox("故障部位", ["全部"] + all_parts, key="u_part")
        avail_partno = sorted(part_map[part_map["分類"] == f_part]["零件編號"].astype(str).unique().tolist()) \
            if f_part != "全部" else []
        f_partno = safe_selectbox("零件料號", ["全部"] + avail_partno, key="u_partno",
                                   disabled=(f_part == "全部"))
        f_branch = r2[3].selectbox("維修分公司別", ["全部"] + all_branches, key="u_branch")

        result = usage_table(usage_df, start_ym, end_ym, f_cat, f_io, f_model,
                              f_part, f_partno, f_branch)
        st.subheader(f"查詢結果（{len(result)} 筆）")
        st.dataframe(result, use_container_width=True)

    # ---------------- 故障率 ----------------
    with tab_rate:
        st.subheader("查詢條件")
        r1 = st.columns(4)
        end_ym_r = r1[0].text_input("查找年月（YYYY-MM，代表 2021/01 ~ 該月）",
                                     value="2026-07", key="r_month")
        f_cat_r = r1[1].selectbox("類別", ["全部"] + all_cats, key="r_cat")
        f_io_r = r1[2].selectbox("室內/外機", ["全部"] + all_ios, key="r_io")

        model_scope_r = model_map
        if f_cat_r != "全部": model_scope_r = model_scope_r[model_scope_r["類別"] == f_cat_r]
        if f_io_r != "全部": model_scope_r = model_scope_r[model_scope_r["內外機"] == f_io_r]
        avail_models_r = sorted(model_scope_r["機型"].unique().tolist())
        f_model_r = safe_selectbox("機型", ["全部"] + avail_models_r, key="r_model")

        r2 = st.columns(3)
        f_part_r = r2[0].selectbox("故障部位", ["全部"] + all_parts, key="r_part")
        f_branch_r = r2[1].selectbox("維修分公司別", ["全部"] + all_branches, key="r_branch")

        partno_locked = not (f_cat_r and f_io_r and f_model_r and f_part_r and f_branch_r) or f_part_r == "全部"
        avail_partno_r = sorted(part_map[part_map["分類"] == f_part_r]["零件編號"].astype(str).unique().tolist()) \
            if (not partno_locked and f_part_r != "全部") else []
        f_partno_r = safe_selectbox(
            "零件料號（需前面條件都已選擇）", ["全部"] + avail_partno_r,
            key="r_partno", disabled=partno_locked,
        )
        if partno_locked:
            st.caption("🔒 零件料號需要前面所有條件都選了才能開放；留空 = 該機型所有零件加總")

        result_r = rate_table(usage_df, shipment_df, year_cols, end_ym_r,
                               f_cat_r, f_io_r, f_model_r, f_part_r, f_partno_r, f_branch_r)
        st.subheader(f"查詢結果（{len(result_r)} 筆）")
        st.dataframe(result_r, use_container_width=True)
        st.caption(
            "⚠️ 累積出貨量目前只能抓到「年度」粒度（出貨資料是每年一欄），"
            "所以累積故障率是用『查找年月所在年度以前的完整年份 + 出貨資料裡若有涵蓋到當年度的欄位』相加，"
            "不是精確到月的出貨量。這點需要跟你確認是否可接受，或提供月粒度的出貨資料。"
        )


if __name__ == "__main__":
    main()
