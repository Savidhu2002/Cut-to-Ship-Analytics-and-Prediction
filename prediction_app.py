import joblib
import numpy as np
import pandas as pd
import streamlit as st

CUT_MODEL_PATH = "direct_cut_qty_model.pkl"
SHIP_MODEL_PATH = "direct_ship_qty_model.pkl"
META_PATH = "direct_deploy_meta.pkl"

cut_pipe = joblib.load(CUT_MODEL_PATH)
ship_pipe = joblib.load(SHIP_MODEL_PATH)
meta = joblib.load(META_PATH)

BASE_INPUT_COLS = meta["base_input_cols"]
UI_COLS = meta["ui_cols"]
FEATURE_COLS = meta["feature_cols"]
ALLOWED_VALUES = meta["allowed_values"]
FREQ_MAPS = meta["freq_maps"]
LOOKUPS = meta["lookups"]

CUT_MAPE_PCT = float(meta.get("cut_mape_pct", 5.0))
SHIP_MAPE_PCT = float(meta.get("ship_mape_pct", 5.0))

def clean_text_value(x):
    if x is None:
        return "Unknown"
    x = str(x).strip()
    if x == "" or x.lower() in ["nan", "none", "n/a"]:
        return "Unknown"
    return x

def safe_float_value(x, default=0.0):
    if x is None:
        return default
    s = str(x).strip()
    if s == "":
        return default
    try:
        return float(s.replace(",", ""))
    except Exception:
        return default

def safe_lookup_freq(freq_map, key):
    return float(freq_map.get(key, 0))

def lookup_behavior(row_dict, lookups):
    feat_cols = lookups["feat_cols"]

    for table_name, key_name in [("full", "key_cols"), ("backoff1", "backoff1_keys"), ("backoff2", "backoff2_keys")]:
        table = lookups[table_name]
        keys = lookups[key_name]

        mask = np.ones(len(table), dtype=bool)
        for k in keys:
            mask &= (table[k].astype(str).values == str(row_dict[k]))

        match = table.loc[mask]
        if len(match) > 0:
            return {f: float(match.iloc[0][f]) for f in feat_cols}

    return {f: float(lookups["global_mean"].get(f, 0.0)) for f in feat_cols}

def pct_text(x):
    return f"{x * 100:.1f}%"

def metric_card(label, value, accent_class="blue"):
    return f"""
    <div class="metric-card {accent_class}">
        <div class="metric-label">{label}</div>
        <div class="metric-value">{value}</div>
    </div>
    """

st.set_page_config(page_title="Cut & Ship Prediction Model", page_icon="📊", layout="wide")

st.markdown("""
<style>
    .stApp { background: linear-gradient(180deg, #f4f7fb 0%, #eef3f8 100%); }
    .block-container { padding-top: 2rem; padding-bottom: 2rem; max-width: 1250px; }
    .hero-box { background: linear-gradient(135deg, #0f2747 0%, #163a63 60%, #1d4c7f 100%); padding: 28px 32px; border-radius: 20px; color: white; margin-bottom: 1.2rem; }
    .hero-title { font-size: 2.1rem; font-weight: 700; margin-bottom: 0.2rem; }
    .section-box { background: #ffffff; border-radius: 18px; padding: 20px 22px 16px 22px; box-shadow: 0 10px 24px rgba(31, 45, 61, 0.08); border: 1px solid #dbe5f0; margin-bottom: 1.2rem; }
    .section-title { font-size: 1.2rem; font-weight: 700; color: #102a43; margin-bottom: 0.8rem; }
    .stButton > button { background: linear-gradient(135deg, #0f4c81 0%, #166088 100%); color: white; border: none; border-radius: 12px; padding: 0.7rem 1.6rem; font-weight: 700; }
    .metric-card { background: #ffffff; border-radius: 18px; padding: 18px 20px; border: 1px solid #dbe5f0; min-height: 118px; margin-bottom: 14px; position: relative; overflow: hidden; }
    .metric-card::before { content: ""; position: absolute; top: 0; left: 0; width: 100%; height: 5px; }
    .metric-card.blue::before { background: linear-gradient(90deg, #0f4c81, #1b6ca8); }
    .metric-card.teal::before { background: linear-gradient(90deg, #0d6e6e, #1b9aaa); }
    .metric-card.gold::before { background: linear-gradient(90deg, #b8860b, #d4a017); }
    .metric-label { font-size: 0.95rem; font-weight: 600; color: #486581; margin-bottom: 0.55rem; }
    .metric-value { font-size: 2rem; font-weight: 800; color: #102a43; line-height: 1.1; }
    .results-title { font-size: 1.35rem; font-weight: 700; color: #102a43; margin-bottom: 0.85rem; }
    .range-box { background: #f8fbff; border: 1px solid #dbe5f0; border-radius: 14px; padding: 14px 16px; margin-top: 8px; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero-box">
    <div class="hero-title">Cut &amp; Ship Prediction Model</div>
</div>
""", unsafe_allow_html=True)

st.markdown('<div class="section-box">', unsafe_allow_html=True)
st.markdown('<div class="section-title">Enter Order Details</div>', unsafe_allow_html=True)

cols = st.columns(3)
inputs = {}

for i, col in enumerate(UI_COLS):
    with cols[i % 3]:
        if col == "Order Qty":
            inputs[col] = st.number_input(col, min_value=1.0, step=1.0, value=1000.0)
        elif col == "Pcs":
            raw_val = st.text_input(col, value="", placeholder="")
            inputs[col] = safe_float_value(raw_val, default=0.0)
        else:
            options = [""] + ALLOWED_VALUES.get(col, ["Unknown"])
            selected = st.selectbox(col, options=options)
            inputs[col] = selected if str(selected).strip() != "" else "Unknown"

predict_clicked = st.button("Predict")
st.markdown('</div>', unsafe_allow_html=True)

if predict_clicked:
    try:
        eps = 1e-9
        order_qty = float(inputs["Order Qty"])

        row = {}
        for col in BASE_INPUT_COLS:
            if col == "Order Qty":
                row[col] = float(inputs[col])
            elif col == "Pcs":
                row[col] = safe_float_value(inputs[col], default=0.0)
            else:
                row[col] = clean_text_value(inputs[col])

        row["Year_Month"] = f"{row['Year']}_{row['Month']}"
        row["Div_Unit"] = f"{row['Div']}_{row['Unit']}"
        row["Season_Garment"] = f"{row['Season']}_{row['Garment item type']}"
        row["CallingName_Garment"] = f"{row['Calling Name']}_{row['Garment item type']}"
        row["Operation_Type"] = f"{row['Operation']}_{row['Type']}"
        row["Operation_Operation2"] = f"{row['Operation']}_{row['Operation 2']}"
        row["Pcs_per_OrderQty"] = row["Pcs"] / (row["Order Qty"] + eps)

        row["Reason_Count_NonZero"] = 0.0
        row["Total_Reason_Qty"] = 0.0
        row["Damage_Total"] = 0.0
        row["Transfer_Total"] = 0.0
        row["Sample_Total"] = 0.0
        row["Quality_Total"] = 0.0
        row["Reconciliation_Total"] = 0.0
        row["Has_Any_Reason"] = 0
        row["Has_Transfer"] = 0
        row["Has_Damage"] = 0
        row["Has_Reconciliation_Issue"] = 0

        row["Year_Freq"] = safe_lookup_freq(FREQ_MAPS.get("Year_Freq", {}), row["Year"])
        row["Calling_Name_Freq"] = safe_lookup_freq(FREQ_MAPS.get("Calling_Name_Freq", {}), row["Calling Name"])
        row["Garment_Type_Freq"] = safe_lookup_freq(FREQ_MAPS.get("Garment_Type_Freq", {}), row["Garment item type"])
        row["Div_Freq"] = safe_lookup_freq(FREQ_MAPS.get("Div_Freq", {}), row["Div"])
        row["Unit_Freq"] = safe_lookup_freq(FREQ_MAPS.get("Unit_Freq", {}), row["Unit"])
        row["Operation_Freq"] = safe_lookup_freq(FREQ_MAPS.get("Operation_Freq", {}), row["Operation"])

        lookup_row = {
            "Year": row["Year"],
            "Calling Name": row["Calling Name"],
            "Div": row["Div"],
            "Season": row["Season"],
            "Garment item type": row["Garment item type"],
            "Unit": row["Unit"],
            "Operation": row["Operation"],
            "Month": row["Month"],
            "Type": row["Type"],
            "Operation 2": row["Operation 2"],
        }

        hist = lookup_behavior(lookup_row, LOOKUPS)
        row["Hist_Cut Qty"] = hist.get("Cut Qty", 0.0)
        row["Hist_Ship Qty"] = hist.get("Ship Qty", 0.0)
        row["Hist_Order Qty"] = hist.get("Order Qty", 0.0)
        row["Hist_Pcs_per_OrderQty"] = hist.get("Pcs_per_OrderQty", 0.0)
        row["Hist_Total_Reason_Qty"] = hist.get("Total_Reason_Qty", 0.0)
        row["Hist_Damage_Total"] = hist.get("Damage_Total", 0.0)
        row["Hist_Transfer_Total"] = hist.get("Transfer_Total", 0.0)
        row["Hist_Quality_Total"] = hist.get("Quality_Total", 0.0)

        numeric_defaults = {
            "Pcs", "Order Qty", "Pcs_per_OrderQty",
            "Year_Freq", "Calling_Name_Freq", "Garment_Type_Freq", "Div_Freq", "Unit_Freq", "Operation_Freq",
            "Reason_Count_NonZero", "Total_Reason_Qty", "Damage_Total", "Transfer_Total", "Sample_Total",
            "Quality_Total", "Reconciliation_Total", "Has_Any_Reason", "Has_Transfer", "Has_Damage",
            "Has_Reconciliation_Issue", "Hist_Cut Qty", "Hist_Ship Qty", "Hist_Order Qty", "Hist_Pcs_per_OrderQty",
            "Hist_Total_Reason_Qty", "Hist_Damage_Total", "Hist_Transfer_Total", "Hist_Quality_Total"
        }

        for col in FEATURE_COLS:
            if col not in row:
                row[col] = 0.0 if col in numeric_defaults else "Unknown"

        X = pd.DataFrame([row])[FEATURE_COLS].copy()

        cut_qty_pred = float(cut_pipe.predict(X)[0])
        ship_qty_pred = float(ship_pipe.predict(X)[0])

        cut_qty_pred = max(0.0, cut_qty_pred)
        ship_qty_pred = max(0.0, ship_qty_pred)

        cut_low = cut_qty_pred * (1 - CUT_MAPE_PCT / 100.0)
        cut_high = cut_qty_pred * (1 + CUT_MAPE_PCT / 100.0)
        ship_low = ship_qty_pred * (1 - SHIP_MAPE_PCT / 100.0)
        ship_high = ship_qty_pred * (1 + SHIP_MAPE_PCT / 100.0)

        cut_ship = ship_qty_pred / (cut_qty_pred + eps)
        order_ship = ship_qty_pred / (order_qty + eps)
        order_cut = cut_qty_pred / (order_qty + eps)

        st.markdown('<div class="section-box">', unsafe_allow_html=True)
        st.markdown('<div class="results-title">Prediction Results</div>', unsafe_allow_html=True)

        r1, r2, r3 = st.columns(3)
        with r1:
            st.markdown(metric_card("Predicted Cut Qty", f"{int(round(cut_qty_pred)):,}", "blue"), unsafe_allow_html=True)
        with r2:
            st.markdown(metric_card("Predicted Ship Qty", f"{int(round(ship_qty_pred)):,}", "teal"), unsafe_allow_html=True)
        with r3:
            st.markdown(metric_card("Order Qty", f"{int(round(order_qty)):,}", "gold"), unsafe_allow_html=True)

        r4, r5, r6 = st.columns(3)
        with r4:
            st.markdown(metric_card("Cut / Ship", pct_text(cut_ship), "blue"), unsafe_allow_html=True)
        with r5:
            st.markdown(metric_card("Order / Ship", pct_text(order_ship), "teal"), unsafe_allow_html=True)
        with r6:
            st.markdown(metric_card("Order / Cut", pct_text(order_cut), "gold"), unsafe_allow_html=True)

        st.markdown('<div class="range-box">', unsafe_allow_html=True)
        st.markdown("**Estimated Prediction Range Based on Historical Model Error**")
        st.markdown(
            f"- **Cut Qty Range:** {int(round(cut_low)):,} to {int(round(cut_high)):,} (±{CUT_MAPE_PCT:.1f}%)\n"
            f"- **Ship Qty Range:** {int(round(ship_low)):,} to {int(round(ship_high)):,} (±{SHIP_MAPE_PCT:.1f}%)"
        )
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('</div>', unsafe_allow_html=True)

    except Exception as e:
        st.error(f"Prediction failed: {e}")