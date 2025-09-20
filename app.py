
import os
import re
import streamlit as st
import pandas as pd
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Tuple, Optional

st.set_page_config(page_title="Purpose Built – Fulfillment Extractor", page_icon="📦", layout="wide")

APP_DIR = Path(__file__).parent
STORE_XLS_PATH = APP_DIR / "data" / "store_master_data.xls"   # bundled store master

# -------- Final ShipStation columns --------
EXPECTED_COLS = [
    "First Name","Last Name","Quantity","Style #","Item Description","Color","SIZE",
    "Ship To Attention","Ship To ADDRESS LINE 1","Ship To  ADDRESS LINE 2","Ship To CITY",
    "Ship To STATE","Ship To POSTAL CODE","Third Party Shipping Account","Shipper/Ship Method"
]

# -------- Aliases in order blocks --------
ALIASES = {
    "Style #":         ["Style #","Sku","SKU","Style","Item #","Item No","ItemNo","Item"],
    "Item Description":["Item Description","Item","Description","Desc"],
    "Color":           ["Color","Colour"],
    "Store #":         ["Store #","Store Number","Store No","Store","StoreCode","Store ID","StoreID","StoreNum","StoreNumber","Store_No"],
}

# Alpha sizes (numeric sizes are detected dynamically)
ALPHA_SIZE_TOKENS = {"xs","s","m","l","xl","2xl","3xl","4xl","5xl","os","onesize","one size","osfm"}

# -------- Shipping field aliases from master --------
MASTER_ALIAS_SHIP = {
    "Ship To Attention":            ["Ship To Attention","Attention","Store Name","ShipToAttention","Ship To Name"],
    "Ship To ADDRESS LINE 1":       ["Ship To ADDRESS LINE 1","Ship To Address Line 1","Address1","Address 1","ShipToAddress1"],
    "Ship To  ADDRESS LINE 2":      ["Ship To  ADDRESS LINE 2","Ship To Address Line 2","Address2","Address 2","ShipToAddress2"],
    "Ship To CITY":                 ["Ship To CITY","Ship To City","City","ShipToCity"],
    "Ship To STATE":                ["Ship To STATE","Ship To State","State","ShipToState"],
    "Ship To POSTAL CODE":          ["Ship To POSTAL CODE","Ship To Postal Code","Postal Code","Zip","Zip Code","Postcode","ShipToPostalCode"],
    "Third Party Shipping Account": ["Third Party Shipping Account","3rd Party Account","ThirdPartyAccount"],
    "Shipper/Ship Method":          ["Shipper/Ship Method","Ship Method","Carrier Service","ShipMethod"],
}

# ============== helpers =================
def _norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(s).strip().lower())

def find_col_by_alias(aliases: List[str], columns: List[str]) -> Optional[str]:
    lut = {_norm(c): c for c in columns}
    for a in aliases:
        key = _norm(a)
        if key in lut:
            return lut[key]
    return None

def normalize_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def normalize_store_key(series: pd.Series) -> pd.Series:
    """Normalize store codes: 1, '1', ' 1 ', '#1', '1.0' -> '1'."""
    s = series.astype(str).str.strip()
    s = s.str.replace(r"\.0$", "", regex=True)
    digits = s.str.extract(r"(\d+)", expand=False)
    s = digits.fillna(s)
    return s

def is_size_token(token: str) -> bool:
    """Treat a header cell as a size if it's an alpha size or a clean numeric like 8, 10, 12, 14."""
    if token is None:
        return False
    t = str(token).strip()
    if not t or t.lower() == "none":
        return False
    tn = _norm(t)
    if tn in ALPHA_SIZE_TOKENS:
        return True
    return bool(re.fullmatch(r"\d+[A-Za-z]*", t))

# Cache master; refresh when file changes
@st.cache_data(show_spinner=False)
def load_store_master(mtime: float) -> pd.DataFrame:
    df = pd.read_excel(STORE_XLS_PATH, dtype=str)  # requires xlrd>=2.0.1 for .xls
    return normalize_cols(df)

def row_has_header_signature(values: List[str]) -> bool:
    """Detect a header row by Style/SKU, Item Description, Color, Store # presence."""
    vals = {_norm(v) for v in values if isinstance(v, str)}
    has_sku   = any(_norm(a) in vals for a in ALIASES["Style #"])
    has_desc  = any(_norm(a) in vals for a in ALIASES["Item Description"])
    has_color = any(_norm(a) in vals for a in ALIASES["Color"])
    has_store = any(_norm(a) in vals for a in ALIASES["Store #"])
    return has_sku and has_desc and has_color and has_store

def detect_header_rows(raw: pd.DataFrame) -> List[int]:
    header_idxs = []
    for i in range(len(raw)):
        row = [str(x) if pd.notna(x) else "" for x in raw.iloc[i].tolist()]
        if row_has_header_signature(row):
            header_idxs.append(i)
    return header_idxs

def parse_blocks_with_sizes(raw: pd.DataFrame) -> Tuple[pd.DataFrame, List[int]]:
    """
    Header pattern per block:
      ... [Style/SKU][Item Description][Color][Store #][Allocation %][size columns...]
    Sizes may be alpha (XS..5XL/OS) or numeric (8,10,12,14,...). Allocation is ignored.
    """
    df = raw.copy()
    header_rows = detect_header_rows(df)
    items = []

    for idx, h in enumerate(header_rows):
        start = h + 1
        end   = header_rows[idx + 1] if idx + 1 < len(header_rows) else len(df)
        header_vals = [str(x) if pd.notna(x) else "" for x in df.iloc[h].tolist()]

        cols = header_vals
        lut = {_norm(c): j for j, c in enumerate(cols) if c}

        def idx_for(aliases: List[str]) -> Optional[int]:
            for a in aliases:
                k = _norm(a)
                if k in lut:
                    return lut[k]
            return None

        sku_i   = idx_for(ALIASES["Style #"])
        item_i  = idx_for(ALIASES["Item Description"])
        color_i = idx_for(ALIASES["Color"])
        store_i = idx_for(ALIASES["Store #"])
        if sku_i is None or item_i is None or color_i is None or store_i is None:
            continue

        alloc_i = store_i + 1  # (ignored)
        # contiguous size columns after allocation
        size_idxs, size_names = [], []
        started = False
        for j in range(alloc_i + 1, len(cols)):
            token = cols[j]
            if is_size_token(token):
                size_idxs.append(j)
                size_names.append(str(token).strip())
                started = True
            else:
                if started:
                    break

        if not size_idxs:
            continue

        block = df.iloc[start:end].values.tolist()
        for row in block:
            sku   = str(row[sku_i]).strip()   if sku_i   < len(row) and pd.notna(row[sku_i])   else ""
            desc  = str(row[item_i]).strip()  if item_i  < len(row) and pd.notna(row[item_i])  else ""
            color = str(row[color_i]).strip() if color_i < len(row) and pd.notna(row[color_i]) else ""
            store = str(row[store_i]).strip() if store_i < len(row) and pd.notna(row[store_i]) else ""
            if not sku or not store:
                continue

            for j, sz_name in zip(size_idxs, size_names):
                qty = 0.0
                if j < len(row) and pd.notna(row[j]):
                    try:
                        qty = float(str(row[j]).strip())
                    except:
                        qty = 0.0
                if qty > 0:
                    items.append({
                        "Style #": sku,
                        "Item Description": desc,
                        "Color": color,
                        "SIZE": sz_name,
                        "Quantity": int(qty) if float(qty).is_integer() else qty,
                        "Store #": store
                    })

    parsed = pd.DataFrame(items, columns=["Style #","Item Description","Color","SIZE","Quantity","Store #"])
    return normalize_cols(parsed), header_rows

def pick_master_shipping_cols(columns: List[str]) -> Dict[str, Optional[str]]:
    pick = {}
    for target, aliases in MASTER_ALIAS_SHIP.items():
        pick[target] = find_col_by_alias(aliases, columns)
    return pick

def build_final_output(merged: pd.DataFrame, ship_pick: Dict[str, Optional[str]]) -> pd.DataFrame:
    out = pd.DataFrame(columns=EXPECTED_COLS)

    # Line-item fields
    for c in ["Quantity","Style #","Item Description","Color","SIZE"]:
        out[c] = merged[c].astype(str) if c in merged.columns else ""

    # Shipping from master (now in merged)
    for tgt in ["Ship To Attention","Ship To ADDRESS LINE 1","Ship To  ADDRESS LINE 2","Ship To CITY",
                "Ship To STATE","Ship To POSTAL CODE","Third Party Shipping Account","Shipper/Ship Method"]:
        src = ship_pick.get(tgt)
        out[tgt] = merged[src].astype(str) if src and src in merged.columns else ""

    # --- FORCE constant names on every row (cannot be overridden) ---
    out.loc[:, "First Name"] = "Purpose"
    out.loc[:, "Last Name"]  = "Built"

    out = out.applymap(lambda v: v.strip() if isinstance(v, str) else v)
    if "Ship To Attention" in out.columns:
        out = out.sort_values(by=["Ship To Attention","Ship To STATE","Ship To CITY","Ship To ADDRESS LINE 1"], na_position="last")
    return out

# ------ robust guessing for the master/order "store number" column ------
def guess_store_column(df: pd.DataFrame) -> Optional[str]:
    if df is None or df.empty:
        return None
    cols = list(df.columns)
    def score(cname: str) -> float:
        s = normalize_store_key(df[cname])
        nonnull = s[s.notna() & (s.str.len() > 0)]
        if len(nonnull) == 0:
            return 0.0
        frac_numeric = (nonnull.str.fullmatch(r"\d+").fillna(False)).mean()
        bonus = 0.0
        cn = _norm(cname)
        if "store" in cn: bonus += 0.5
        if "number" in cn or cn.endswith("no") or "id" in cn: bonus += 0.25
        return frac_numeric + bonus
    candidates = sorted(cols, key=score, reverse=True)
    best = candidates[0] if candidates else None
    if best:
        s = normalize_store_key(df[best])
        nonnull = s[s.notna() & (s.str.len() > 0)]
        if len(nonnull) and (nonnull.str.fullmatch(r"\d+").fillna(False)).mean() >= 0.6:
            return best
    return None

def robust_auto_merge(order_df: pd.DataFrame, store_master: pd.DataFrame) -> Tuple[pd.DataFrame, str, str]:
    """Auto-merge on store number with strong normalization; returns merged and the keys used."""
    order_key  = find_col_by_alias(ALIASES["Store #"],  list(order_df.columns))
    master_key = find_col_by_alias(ALIASES["Store #"],  list(store_master.columns))
    if master_key is None:
        master_key = guess_store_column(store_master)
    if order_key is None:
        order_key = guess_store_column(order_df)
    if not order_key or not master_key:
        raise ValueError(
            f"Could not find join keys.\nOrder cols: {list(order_df.columns)}\nMaster cols: {list(store_master.columns)}"
        )
    odf = order_df.copy()
    mdf = store_master.copy()
    odf["_join_key"] = normalize_store_key(odf[order_key])
    mdf["_join_key"] = normalize_store_key(mdf[master_key])
    merged = odf.merge(mdf, on="_join_key", how="left", suffixes=("", "_store"))
    merged.drop(columns=["_join_key"], inplace=True)
    return normalize_cols(merged), order_key, master_key

# ================= UI =================
st.title("📦 Purpose Built – Fulfillment Extractor")
st.caption("Upload the **new order tab** (.xlsx/.csv). The app detects header rows, explodes alpha/numeric sizes, "
           "auto-merges with the bundled Store Master on Store Number, and exports the fulfillment CSV. "
           "First/Last Name are always set to Purpose / Built.")

mtime = os.path.getmtime(STORE_XLS_PATH) if STORE_XLS_PATH.exists() else 0.0
store_master = load_store_master(mtime)

with st.expander("Preview: Store Master (bundled)"):
    st.dataframe(store_master.head(25), use_container_width=True)
    st.caption(f"Master columns: {list(store_master.columns)}  •  File: {STORE_XLS_PATH.name}")

uploaded = st.file_uploader("Upload the 'new order tab' export (.xlsx or .csv)", type=["xlsx","csv"])
raw = None
if uploaded:
    try:
        if uploaded.name.lower().endswith(".csv"):
            raw = pd.read_csv(uploaded, header=None, dtype=str)
        else:
            xls = pd.ExcelFile(uploaded)
            raw = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=None, dtype=str)
        st.success(f"Loaded raw file • shape {raw.shape}")
        st.dataframe(raw.head(30), use_container_width=True)
    except Exception as e:
        st.error(f"Could not read file: {e}")

if raw is not None:
    st.markdown("---")
    st.subheader("Step 1: Detect headers & build line items (alpha or numeric sizes)")
    parsed, header_rows = parse_blocks_with_sizes(raw)
    if not len(parsed):
        st.error("No line items detected. If headers use different words, we can add aliases.")
        st.stop()
    st.info(f"Detected header rows at indices: {header_rows} (0-based).")
    st.dataframe(parsed.head(50), use_container_width=True)

    st.subheader("Step 2: Auto-merge with Store Master (no manual selection)")
    try:
        merged, used_order_key, used_master_key = robust_auto_merge(parsed, store_master)
        st.success(f"Merged rows: {len(merged)}")
        st.caption(f"Join used ➜ Order key: **{used_order_key}**  •  Master key: **{used_master_key}**")
        st.dataframe(merged.head(50), use_container_width=True)
    except Exception as e:
        st.error(str(e))
        st.stop()

    st.subheader("Step 3: Build & export fulfillment CSV")
    ship_pick = pick_master_shipping_cols(list(merged.columns))
    final = build_final_output(merged, ship_pick)
    st.success(f"Ready • Rows: {len(final)} • Columns: {len(final.columns)}")
    st.dataframe(final.head(50), use_container_width=True)

    fname = f"Purpose_Built_Fulfillment_Sheet_{datetime.now().strftime('%Y-%m-%d')}.csv"
    st.download_button("Download CSV", data=final.to_csv(index=False).encode("utf-8"),
                       file_name=fname, mime="text/csv")
