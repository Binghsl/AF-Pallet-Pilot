import streamlit as st
import pandas as pd
import math
from io import BytesIO

st.set_page_config(page_title="📦 Pallet Layer Simulation with Freeform Leftover Mixing", layout="wide")
st.title(":package: Pallet Layer Simulation - Freeform Leftover Mixing")

st.markdown("""
**Layer stacking rules:**  
1) Group those with the same layer dimension and palletize them (no mixing PNs or dimensions for these full pallets).  
2) Only allow mixed loading (even for different PNs and dimensions) for remaining pallet layers ("leftover layers").  
""")

# ============================================================
# Upload helpers (template + parsing)
# ============================================================
UPLOAD_REQUIRED = [
    "Part No.", "MC Qty", "Length(CM)", "Width(CM)", "Height(CM)",
    "Box/Layer", "Max Layer", "MC Weight (gram)"
]

def download_template_bytes() -> bytes:
    df = pd.DataFrame(columns=UPLOAD_REQUIRED)
    bio = BytesIO()
    df.to_excel(bio, index=False)
    bio.seek(0)
    return bio.read()


def normalize_uploaded(df: pd.DataFrame) -> pd.DataFrame:
    """Map uploaded headers to the required template headers; be forgiving with names."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    alias = {
        "part no.": "Part No.", "part no": "Part No.", "pn": "Part No.",
        "mc qty": "MC Qty", "qty": "MC Qty",
        "length(cm)": "Length(CM)", "length (cm)": "Length(CM)", "length": "Length(CM)",
        "width(cm)": "Width(CM)", "width (cm)": "Width(CM)", "width": "Width(CM)",
        "height(cm)": "Height(CM)", "height (cm)": "Height(CM)", "height": "Height(CM)",
        "box/layer": "Box/Layer", "box per layer": "Box/Layer",
        "max layer": "Max Layer", "max_layer": "Max Layer",
        "mc weight (gram)": "MC Weight (gram)", "weight (gram)": "MC Weight (gram)", "weight": "MC Weight (gram)",
    }

    lower_to_orig = {c.lower(): c for c in df.columns}
    renames = {}
    for lc, orig in lower_to_orig.items():
        if lc in alias:
            renames[orig] = alias[lc]
    df = df.rename(columns=renames)

    # Ensure required columns exist
    for col in UPLOAD_REQUIRED:
        if col not in df.columns:
            df[col] = pd.NA

    # Keep required & order
    df = df[UPLOAD_REQUIRED]

    # Types
    for c in ["MC Qty", "Length(CM)", "Width(CM)", "Height(CM)", "Box/Layer", "Max Layer", "MC Weight (gram)"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    df["Part No."] = df["Part No."].astype(str).str.strip()
    df = df.dropna(subset=["Part No."])
    df = df[df["Part No."] != ""].reset_index(drop=True)
    return df


def to_internal_schema(df: pd.DataFrame) -> pd.DataFrame:
    """Convert template columns to internal names used by the logic."""
    out = df.rename(columns={
        "Part No.": "Part No",
        "MC Qty": "Quantity",
        "Length(CM)": "Length (cm)",
        "Width(CM)": "Width (cm)",
        "Height(CM)": "Height (cm)",
    })
    return out

# ============================================================
# Defaults (used if no upload)
# ============================================================
default_data = [
    {"Part No": "51700", "Length (cm)": 60, "Width (cm)": 29, "Height (cm)": 29, "Quantity": 14, "Box/Layer": 6, "Max Layer": 4, "MC Weight (gram)": 0},
    {"Part No": "52363", "Length (cm)": 54, "Width (cm)": 38, "Height (cm)": 31, "Quantity": 5, "Box/Layer": 5, "Max Layer": 4, "MC Weight (gram)": 0},
    {"Part No": "61385", "Length (cm)": 51, "Width (cm)": 35, "Height (cm)": 30, "Quantity": 78, "Box/Layer": 6, "Max Layer": 4, "MC Weight (gram)": 0},
    {"Part No": "61386", "Length (cm)": 41, "Width (cm)": 35, "Height (cm)": 30, "Quantity": 52, "Box/Layer": 8, "Max Layer": 4, "MC Weight (gram)": 0},
    {"Part No": "61387", "Length (cm)": 41, "Width (cm)": 35, "Height (cm)": 30, "Quantity": 18, "Box/Layer": 8, "Max Layer": 4, "MC Weight (gram)": 0},
    {"Part No": "61388", "Length (cm)": 41, "Width (cm)": 35, "Height (cm)": 30, "Quantity": 52, "Box/Layer": 8, "Max Layer": 4, "MC Weight (gram)": 0},
    {"Part No": "61400", "Length (cm)": 50, "Width (cm)": 30, "Height (cm)": 25, "Quantity": 40, "Box/Layer": 5, "Max Layer": 5, "MC Weight (gram)": 0},
    {"Part No": "61401", "Length (cm)": 55, "Width (cm)": 32, "Height (cm)": 28, "Quantity": 35, "Box/Layer": 5, "Max Layer": 5, "MC Weight (gram)": 0},
    {"Part No": "61402", "Length (cm)": 48, "Width (cm)": 34, "Height (cm)": 27, "Quantity": 36, "Box/Layer": 6, "Max Layer": 4, "MC Weight (gram)": 0},
    {"Part No": "61403", "Length (cm)": 45, "Width (cm)": 30, "Height (cm)": 25, "Quantity": 24, "Box/Layer": 4, "Max Layer": 4, "MC Weight (gram)": 0},
]

# ============================================================
# Data source: upload or manual
# ============================================================
st.header("Data Source")

c1, c2 = st.columns([1,3])
with c1:
    st.download_button(
        "Download Upload Template (.xlsx)",
        data=download_template_bytes(),
        file_name="boxes_upload_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
with c2:
    uploaded = st.file_uploader("Upload Excel/CSV in the given template", type=["xlsx", "xls", "csv"])

# PN limit (max 10)
max_pn_limit = 10
pn_count = st.slider("Number of Part Numbers (PNs) to simulate", 1, max_pn_limit, 10)

# Build the working DataFrame
if uploaded is not None:
    try:
        if uploaded.name.lower().endswith(".csv"):
            raw_df = pd.read_csv(uploaded)
        else:
            raw_df = pd.read_excel(uploaded)
        raw_df = normalize_uploaded(raw_df)
        in_df = to_internal_schema(raw_df)
        # Coerce numeric types & take first pn_count rows for initial display
        for c in ["Length (cm)", "Width (cm)", "Height (cm)", "Quantity", "Box/Layer", "Max Layer", "MC Weight (gram)"]:
            in_df[c] = pd.to_numeric(in_df[c], errors="coerce")
        in_df = in_df.head(pn_count)
        st.success(f"Loaded {len(in_df)} rows from upload.")
    except Exception as e:
        st.error(f"Failed to parse uploaded file: {e}")
        in_df = pd.DataFrame(default_data[:pn_count])
else:
    in_df = pd.DataFrame(default_data[:pn_count])

# Editable grid
box_df = st.data_editor(
    in_df,
    num_rows="dynamic",
    use_container_width=True,
    key="box_input"
)

# ============================================================
# Pallet settings
# ============================================================
st.header("Pallet Settings")
pallet_length = st.number_input("Pallet Length (cm)", min_value=50.0, value=120.0)
pallet_width  = st.number_input("Pallet Width (cm)",  min_value=50.0, value=100.0)

# NEW: base height adjustable (default 15 cm)
pallet_base_height = st.number_input("Pallet Base Height (cm)", min_value=0.0, value=15.0, step=0.5)

# CHANGED: default max total height from 150 -> 155 cm
pallet_max_total_height = st.number_input("Max Pallet Height Including Base (cm)", min_value=100.0, max_value=200.0, value=155.0)

max_stack_height = pallet_max_total_height - pallet_base_height

# Volumetric divisor (cm³/kg) for volumetric weight calculation
volumetric_divisor = st.number_input(
    "Volumetric divisor (cm³/kg)", min_value=3000, max_value=10000, value=6000, step=500
)

# Pallet tare (actual pallet weight)
pallet_tare_weight_kg = st.number_input(
    "Pallet Tare Weight (kg)", min_value=0.0, value=20.0, step=0.5
)

# ============================================================
# Core functions (keeps partial-layer overcount)
# ============================================================
def explode_layers(df):
    """
    Create one 'layer' record per ceiling(Quantity / Box/Layer).
    IMPORTANT: Boxes in Layer = Box/Layer even for last partial layer (overcount by design).
    """
    layers = []
    for idx, row in df.iterrows():
        try:
            qty_total = int(row["Quantity"])
            per_layer_cap = int(row["Box/Layer"])
        except Exception:
            continue
        if per_layer_cap <= 0:
            continue

        layers_needed = math.ceil(qty_total / per_layer_cap)
        for _ in range(layers_needed):
            layer_boxes = per_layer_cap  # overcount intentionally
            layers.append({
                "Part No": str(row["Part No"]),
                "Box Length": float(row["Length (cm)"]),
                "Box Width": float(row["Width (cm)"]),
                "Box Height": float(row["Height (cm)"]),
                "Box/Layer": per_layer_cap,
                "Max Layer": int(row["Max Layer"]),
                "Boxes in Layer": layer_boxes,
                "Layer Height": float(row["Height (cm)"]),
                "Layer Source": idx,
                "MC Weight (gram)": float(row.get("MC Weight (gram)", 0) or 0),
            })
    return layers


def pack_layers_by_pn_and_dimension(layers):
    pallets = []
    unassigned_layers = []
    df_layers = pd.DataFrame(layers)
    if df_layers.empty:
        return [], []

    group_cols = ["Box Length", "Box Width", "Box Height", "Box/Layer", "Max Layer", "Part No"]
    for key, group in df_layers.groupby(group_cols):
        group = group.copy()
        box_height = key[2]
        max_layer = key[4]
        stacking_height = max_layer * box_height
        if stacking_height > max_stack_height:
            max_layer = math.floor(max_stack_height / box_height) if box_height > 0 else 0

        num_full_pallets = len(group) // max_layer if max_layer > 0 else 0

        for i in range(num_full_pallets):
            these_layers = group.iloc[i*max_layer:(i+1)*max_layer]
            pallet_height = float(these_layers["Layer Height"].sum())
            util = round((pallet_height / (max_layer * box_height) * 100), 1) if max_layer > 0 and box_height > 0 else 0.0
            pallets.append({
                "Pallet Group": "Full (No Mix)",
                "Part Nos": [these_layers["Part No"].iloc[0]],
                "Box Length": key[0],
                "Box Width": key[1],
                "Box Height": key[2],
                "Box/Layer": key[3],
                "Max Layer": max_layer,
                "Pallet Layers": max_layer,
                "Total Boxes": int(these_layers["Boxes in Layer"].sum()),  # overcount kept
                "Pallet Height (cm)": pallet_height,
                "Height Utilization (%)": util,
                "Layer Details": these_layers.to_dict("records")
            })

        # Remainders to leftovers
        remain = len(group) % max_layer if max_layer > 0 else len(group)
        if remain > 0:
            remain_layers = group.iloc[-remain:]
            for _, lrow in remain_layers.iterrows():
                unassigned_layers.append(lrow)

    return pallets, unassigned_layers


def pack_leftover_layers_any_mix(unassigned_layers):
    pallets = []
    if len(unassigned_layers) == 0:
        return pallets
    df_layers = pd.DataFrame(unassigned_layers)
    leftover_layers = df_layers.copy()
    while not leftover_layers.empty:
        batch_layers = []
        cumulative_height = 0.0
        for idx, row in leftover_layers.iterrows():
            if len(batch_layers) >= int(row["Max Layer"]):
                continue
            if cumulative_height + float(row["Layer Height"]) > max_stack_height:
                break
            batch_layers.append(row)
            cumulative_height += float(row["Layer Height"])
        if not batch_layers:
            break
        batch = pd.DataFrame(batch_layers)
        pallet_height = float(batch["Layer Height"].sum())
        all_pns = sorted(batch["Part No"].unique().tolist())
        dim_str = "; ".join(f'{r["Part No"]}:{r["Box Length"]}x{r["Box Width"]}x{r["Box Height"]}' for _, r in batch.iterrows())
        util = 0.0
        if len(batch) > 0:
            # Util vs theoretical cap from the tightest (min Max Layer * height) within the mix but clamped by max_stack_height
            theoretical = min(max_stack_height, min(batch["Max Layer"] * batch["Box Height"]))
            if theoretical > 0:
                util = round((pallet_height / theoretical) * 100, 1)
        pallets.append({
            "Pallet Group": "Consolidated (Free Mix)",
            "Part Nos": all_pns,
            "Box Length": "Mixed",
            "Box Width": "Mixed",
            "Box Height": "Mixed",
            "Box/Layer": "Mixed",
            "Max Layer": int(batch["Max Layer"].max()),
            "Pallet Layers": len(batch),
            "Total Boxes": int(batch["Boxes in Layer"].sum()),  # overcount kept
            "Pallet Height (cm)": pallet_height,
            "Height Utilization (%)": util,
            "Layer Details": batch.to_dict("records"),
            "Layer Summary": dim_str
        })
        leftover_layers = leftover_layers.iloc[len(batch):]
    return pallets


def create_consolidated_csv(pallets, pallet_L, pallet_W, pallet_base_H, vol_divisor, pallet_tare_kg):
    def pn_boxes_from_layers(layer_details):
        # Sum Boxes in Layer by PN (keeps overcount behavior)
        counts = {}
        for r in layer_details:
            pn = str(r["Part No"])
            counts[pn] = counts.get(pn, 0) + int(r["Boxes in Layer"])
        return counts

    rows = []
    for i, p in enumerate(pallets):
        total_height = float(p["Pallet Height (cm)"]) + float(pallet_base_H)  # includes base

        # PN → boxes string
        pn_counts = pn_boxes_from_layers(p.get("Layer Details", []))
        pn_boxes_str = "; ".join(f"{k}: {v}" for k, v in pn_counts.items()) if pn_counts else ""

        # Volumetric weight (kg) = (L*W*H in cm³) / divisor
        volume_cm3 = float(pallet_L) * float(pallet_W) * float(total_height)
        vol_weight_kg = volume_cm3 / float(vol_divisor) if vol_divisor else 0.0

        # Actual cargo weight from layers (overcount-based: Boxes in Layer * MC Weight per box)
        total_mc_weight_g = 0.0
        for r in p.get("Layer Details", []):
            total_mc_weight_g += float(r.get("MC Weight (gram)", 0) or 0) * int(r.get("Boxes in Layer", 0) or 0)
        total_mc_weight_kg = total_mc_weight_g / 1000.0

        # Add pallet tare weight
        total_pallet_weight_kg = total_mc_weight_kg + float(pallet_tare_kg)

        rows.append({
            "Pallet No": i + 1,
            "Pallet Group": p["Pallet Group"],
            "Part Nos": ", ".join(map(str, p["Part Nos"])),
            # Removed box L/W/H from export
            "Pallet Length (cm)": float(pallet_L),
            "Pallet Width (cm)": float(pallet_W),
            "Pallet Height (cm)": round(total_height, 1),  # includes base
            "Pallet Dimension (cm)": f"{int(pallet_L)}x{int(pallet_W)}x{round(total_height)}",
            "Pallet Layers": p["Pallet Layers"],
            "Max Layer": p["Max Layer"],
            "Total Boxes (overcount)": p["Total Boxes"],
            "PN Boxes": pn_boxes_str,
            "Height Utilization (%)": p["Height Utilization (%)"],
            "Layer Summary": p.get("Layer Summary", ""),
            "Volumetric Weight (kg)": round(vol_weight_kg, 2),
            "Actual Pallet Weight (kg)": round(total_pallet_weight_kg, 2)  # NEW
        })
    return pd.DataFrame(rows)

# ============================================================
# Run simulation
# ============================================================
if st.button("Simulate and Consolidate"):
    if box_df.empty:
        st.error("Please enter or upload box data")
    else:
        layers = explode_layers(box_df)
        full_pallets, unassigned_layers = pack_layers_by_pn_and_dimension(layers)
        mixed_pallets = pack_leftover_layers_any_mix(unassigned_layers)
        all_pallets = full_pallets + mixed_pallets
        csv_df = create_consolidated_csv(
            all_pallets,
            pallet_length,
            pallet_width,
            pallet_base_height,
            volumetric_divisor,
            pallet_tare_weight_kg
        )

        st.success(f"Total simulated pallets: {len(csv_df)} (including consolidated)")
        st.download_button(
            label="⬇️ Download Pallet Plan CSV",
            data=csv_df.to_csv(index=False).encode("utf-8"),
            file_name="pallet_simulation_consolidated.csv",
            mime="text/csv"
        )
        st.dataframe(csv_df)
