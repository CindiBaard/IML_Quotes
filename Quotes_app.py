import sys
import streamlit as st
import pandas as pd
from datetime import datetime

# Optional imports for Google Sheets integration:
try:
    import gspread
    from google.oauth2.service_account import Credentials
    HAS_GSPREAD = True
except ImportError:
    HAS_GSPREAD = False

# Optional import for PDF generation:
try:
    from fpdf import FPDF
    HAS_FPDF = True
except ImportError:
    HAS_FPDF = False


# -----------------------------------------------------------------------------
# DIGITAL TRIAL DIELINE MASTER DATA (EXTRACTED FROM REN-FLEX PDF QUOTES)
# -----------------------------------------------------------------------------

DIELINE_SPECS = {
    "Die 173 - 250ml Jar (274.395mm x 44.121mm)": {
        "quote_ref": "10794/01",
        "die_code": "Die 173",
        "jar_name": "250ml Jar",
        "size_mm": "274.395mm x 44.121mm",
        "moq": 25000,
        "unit_price": 0.5162,
        "run_subtotal": 12905.00,
    },
    "Die 174 - 350ml Jar (274.392mm x 60.183mm)": {
        "quote_ref": "10795/01",
        "die_code": "Die 174",
        "jar_name": "350ml Jar",
        "size_mm": "274.392mm x 60.183mm",
        "moq": 18000,
        "unit_price": 0.7082,
        "run_subtotal": 12747.60,
    },
    "Die 175 - 450ml Jar (274.387mm x 82.041mm)": {
        "quote_ref": "10796/01",
        "die_code": "Die 175",
        "jar_name": "450ml Jar",
        "size_mm": "274.387mm x 82.041mm",
        "moq": 13500,
        "unit_price": 0.9650,
        "run_subtotal": 13027.50,
    },
    "Die 176 - 500ml Jar (291.273mm x 84.605mm)": {
        "quote_ref": "10797/01",
        "die_code": "Die 176",
        "jar_name": "500ml Jar",
        "size_mm": "291.273mm x 84.605mm",
        "moq": 13500,
        "unit_price": 0.9650,
        "run_subtotal": 13027.50,
    },
    "Die 189 - 450ml Std Lightweight Jar (275.378mm x 87.066mm)": {
        "quote_ref": "10798/01",
        "die_code": "Die 189",
        "jar_name": "450ml Std Lightweight Jar",
        "size_mm": "275.378mm x 87.066mm",
        "moq": 13500,
        "unit_price": 0.9650,
        "run_subtotal": 13027.50,
    },
    "Die 211 - 500ml Eve Clicks Tub (292.304mm x 81.936mm)": {
        "quote_ref": "10799/01",
        "die_code": "Die 211",
        "jar_name": "500ml Eve Clicks Tub",
        "size_mm": "292.304mm x 81.936mm",
        "moq": 13500,
        "unit_price": 0.9650,
        "run_subtotal": 13027.50,
    },
    "Die 216 - Clicks Labels (292.304mm x 81.936mm)": {
        "quote_ref": "10800/01",
        "die_code": "Die 216",
        "jar_name": "Clicks Labels",
        "size_mm": "292.304mm x 81.936mm",
        "moq": 13500,
        "unit_price": 0.9650,
        "run_subtotal": 13027.50,
    },
    "Die 264 - 250ml Fabupharm (269.908mm x 44.967mm)": {
        "quote_ref": "10801/01",
        "die_code": "Die 264",
        "jar_name": "250ml Fabupharm",
        "size_mm": "269.908mm x 44.967mm",
        "moq": 25000,
        "unit_price": 0.5162,
        "run_subtotal": 12905.00,
    },
}

DIELINE_OPTIONS = list(DIELINE_SPECS.keys())

DEFAULT_MATERIAL_SPECS = (
    "IML White Gloss, 60mic | Printed CMYK front only | "
    "Finishing: Waterbased Varnish, Die Cutting - Digital | Other Material: Primer | "
    "Note: Subject to +/-10% over/under delivery"
)


# -----------------------------------------------------------------------------
# HELPER FUNCTIONS: GOOGLE SHEETS & PDF GENERATION
# -----------------------------------------------------------------------------

def append_to_google_sheets(data: dict):
    if not HAS_GSPREAD:
        st.warning("gspread library is not installed. Skipping Google Sheets integration.")
        return False

    try:
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        
        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=scopes
        )
        client = gspread.authorize(creds)
        
        SHEET_ID = "1b4Io3QvIJNgI29H1mLSdugKV1nOockWHx-dFuh8YUyE"
        sheet = client.open_by_key(SHEET_ID).sheet1

        headers = [
            "Timestamp", "Total Quantity", "No of Variants", "Variant Details",
            "Substrate Colour", "Substrate Type", "Material Specs", "Target Jars",
            "Colour Matching", "Pantone Qty", "Pantone Refs", "Target Qty",
            "Target Refs", "Litho Print", "Digital Print", "Epson Proof Qty",
            "Layflat Supplied", "Trials Required", "Trial Variants", "Trial Qty Per Variant",
            "Net Subtotal", "15% VAT", "Gross Total"
        ]
        
        existing_row1 = sheet.row_values(1)
        if not existing_row1:
            sheet.append_row(headers)

        flat_variants = [
            f"[{run['dieline']}] {v['variant']}: {v['quantity']}"
            for run in data["Dieline Runs"]
            for v in run["variants"]
        ]

        row = [
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            data["Total Quantity"],
            data["Total Variants"],
            str(flat_variants),
            data["Substrate Colour"],
            data["Substrate Type"],
            data["Material Specs"],
            ", ".join(data["Target Jars"]),
            ", ".join(data["Colour Matching"]),
            data["Pantone Qty"],
            data["Pantone Refs"],
            data["Target Qty"],
            data["Target Refs"],
            data["Litho Print"],
            data["Digital Print"],
            data["Epson Proof Qty"],
            data["Layflat Supplied"],
            "Yes (Digital Trial Quote)",
            data["Total Variants"],
            "Split across variants to MOQ",
            data["Net Subtotal"],
            data["VAT (15%)"],
            data["Gross Total"]
        ]

        sheet.append_row(row)
        return True
    except Exception as e:
        st.error(f"Failed to append to Google Sheet: {e}")
        return False


def generate_pdf_quote(data: dict) -> bytes:
    """Generates a clean PDF document for the digital trial quotation using FPDF."""
    if not HAS_FPDF:
        return b"PDF Generation requires 'fpdf2' package."

    class PDF(FPDF):
        def header(self):
            self.set_font("Helvetica", "B", 16)
            self.set_text_color(30, 41, 59)
            self.cell(0, 10, "In-Mould Label (IML) Digital Trial Quotation", border=False, new_x="LMARGIN", new_y="NEXT", align="L")
            self.set_font("Helvetica", "", 9)
            self.set_text_color(100, 116, 139)
            self.cell(0, 5, f"Generated on: {datetime.now().strftime('%d %B %Y')} | Quote Ref: #IML-TRIAL-{datetime.now().strftime('%M%S')}", border=False, new_x="LMARGIN", new_y="NEXT", align="L")
            self.ln(5)

        def footer(self):
            self.set_y(-15)
            self.set_font("Helvetica", "I", 8)
            self.set_text_color(148, 163, 184)
            self.cell(0, 10, f"Page {self.page_no()} | Confidential IML Digital Trial Quote", align="C")

    pdf = PDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)

    def section_heading(title):
        pdf.set_font("Helvetica", "B", 11)
        pdf.set_fill_color(241, 245, 249)
        pdf.set_text_color(15, 23, 42)
        pdf.cell(0, 7, f"  {title}", border=False, new_x="LMARGIN", new_y="NEXT", fill=True)
        pdf.ln(2)

    def key_value(key, value):
        pdf.set_font("Helvetica", "B", 9)
        pdf.set_text_color(71, 85, 105)
        pdf.cell(60, 6, key, border=False)
        pdf.set_font("Helvetica", "", 9)
        pdf.set_text_color(15, 23, 42)
        pdf.multi_cell(0, 6, str(value), border=False)

    # 1. Digital Trial Runs & Variant Allocation
    section_heading("1. Digital Trial Dieline Runs & Variant Allocation")
    key_value("Total Trial Order Qty:", f"{data['Total Quantity']:,} labels")
    key_value("Total Dieline Runs:", len(data["Dieline Runs"]))
    key_value("Total Label Designs:", data["Total Variants"])
    pdf.ln(2)

    for idx, run in enumerate(data["Dieline Runs"], start=1):
        pdf.set_font("Helvetica", "B", 9)
        pdf.set_fill_color(226, 232, 240)
        pdf.cell(
            0, 6,
            f"Run #{idx}: {run['dieline']} | MOQ: {run['moq']:,} @ R {run['unit_price']:.4f} = R {run['run_charge']:,.2f}",
            border=1, new_x="LMARGIN", new_y="NEXT", fill=True
        )
        pdf.set_font("Helvetica", "B", 8)
        pdf.cell(125, 5, "Label Design / Variant Description", border=1)
        pdf.cell(35, 5, "Allocated Qty", border=1, align="R")
        pdf.cell(30, 5, "% of MOQ", border=1, new_x="LMARGIN", new_y="NEXT", align="R")

        pdf.set_font("Helvetica", "", 8)
        for v in run["variants"]:
            v_name = v["variant"] if v["variant"] else "Unnamed Label Design"
            pct = (v["quantity"] / run["moq"] * 100) if run["moq"] > 0 else 0
            pdf.cell(125, 5, v_name[:75], border=1)
            pdf.cell(35, 5, f"{v['quantity']:,}", border=1, align="R")
            pdf.cell(30, 5, f"{pct:.1f}%", border=1, new_x="LMARGIN", new_y="NEXT", align="R")
        pdf.ln(3)

    # 2. Substrate
    section_heading("2. Substrate Specifications")
    key_value("Substrate Colour:", data["Substrate Colour"])
    key_value("Substrate Finish:", data["Substrate Type"])
    key_value("Material Specs:", data["Material Specs"] if data["Material Specs"] else "Standard Specification")
    pdf.ln(4)

    # 3. Jar Compatibility
    section_heading("3. Target Jar Specifications")
    if data["Target Jars"]:
        for jar in data["Target Jars"]:
            pdf.set_font("Helvetica", "", 9)
            pdf.cell(0, 5, f"  - {jar}", border=False, new_x="LMARGIN", new_y="NEXT")
    else:
        pdf.set_font("Helvetica", "I", 9)
        pdf.cell(0, 5, "No specific jar configurations selected.", border=False, new_x="LMARGIN", new_y="NEXT")
    pdf.ln(4)

    # 4. Print & Colour
    section_heading("4. Print & Colour Specifications")
    key_value("Colour Matching Methods:", ", ".join(data["Colour Matching"]) if data["Colour Matching"] else "None specified")
    if "Pantone" in data["Colour Matching"]:
        key_value("Pantone Labels Qty:", data["Pantone Qty"])
        key_value("Pantone References:", data["Pantone Refs"])
    if "Match to target supplied" in data["Colour Matching"]:
        key_value("Target Match Qty:", data["Target Qty"])
        key_value("Target References:", data["Target Refs"])
    
    key_value("Litho Print Required:", data["Litho Print"])
    key_value("Digital Print Required:", data["Digital Print"])
    key_value("Epson Proofs Required:", f"{data['Epson Proof Qty']} proof(s)")
    key_value("Layflat Supplied:", data["Layflat Supplied"])
    pdf.ln(4)

    # 5. Financial Summary Table
    section_heading("5. Financial Summary")
    
    pdf.set_font("Helvetica", "B", 8)
    pdf.set_fill_color(226, 232, 240)
    pdf.cell(95, 6, "Task Description", border=1, fill=True)
    pdf.cell(35, 6, "Rate / Unit", border=1, align="R", fill=True)
    pdf.cell(25, 6, "Qty", border=1, align="R", fill=True)
    pdf.cell(35, 6, "Sub Total", border=1, new_x="LMARGIN", new_y="NEXT", align="R", fill=True)

    pdf.set_font("Helvetica", "", 8)
    for run in data["Dieline Runs"]:
        desc = f"Digital Trial Run: {run['die_code']} ({run['jar_name']})"
        pdf.cell(95, 6, desc[:55], border=1)
        pdf.cell(35, 6, f"R {run['unit_price']:.4f}", border=1, align="R")
        pdf.cell(25, 6, f"{run['moq']:,}", border=1, align="R")
        pdf.cell(35, 6, f"R {run['run_charge']:,.2f}", border=1, new_x="LMARGIN", new_y="NEXT", align="R")

    extra_items = [
        ("Colour Matching", data.get("cm_net_val", 0.0), data.get("cm_quantity", 0)),
        ("Layflat", data.get("lf_net_val", 0.0), data.get("lf_quantity", 0)),
        ("Additional Setup / Tooling", data.get("pt_net_val", 0.0), data.get("pt_quantity", 0)),
    ]

    for item_desc, net_val, qty in extra_items:
        subtotal = net_val * qty
        pdf.cell(95, 6, item_desc, border=1)
        pdf.cell(35, 6, f"R {net_val:,.2f}", border=1, align="R")
        pdf.cell(25, 6, str(qty), border=1, align="R")
        pdf.cell(35, 6, f"R {subtotal:,.2f}", border=1, new_x="LMARGIN", new_y="NEXT", align="R")

    pdf.ln(2)
    pdf.set_font("Helvetica", "B", 9)
    pdf.cell(155, 6, "Net Subtotal:", border=0, align="R")
    pdf.cell(35, 6, f"R {data['Net Subtotal']:,.2f}", border=1, new_x="LMARGIN", new_y="NEXT", align="R")

    pdf.cell(155, 6, "15% VAT:", border=0, align="R")
    pdf.cell(35, 6, f"R {data['VAT (15%)']:,.2f}", border=1, new_x="LMARGIN", new_y="NEXT", align="R")

    pdf.set_fill_color(241, 245, 249)
    pdf.cell(155, 7, "Gross Total:", border=0, align="R")
    pdf.cell(35, 7, f"R {data['Gross Total']:,.2f}", border=1, new_x="LMARGIN", new_y="NEXT", align="R", fill=True)

    return bytes(pdf.output())


# -----------------------------------------------------------------------------
# SESSION STATE HELPERS & CALLBACKS
# -----------------------------------------------------------------------------

def distribute_moq_evenly(run_idx: int):
    """Distributes the dieline MOQ across the configured number of variants for that run."""
    die_key = st.session_state.get(f"run_die_{run_idx}", DIELINE_OPTIONS[0])
    moq = DIELINE_SPECS[die_key]["moq"]
    n_vars = int(st.session_state.get(f"run_n_vars_{run_idx}", 1))

    base_share = moq // n_vars
    remainder = moq % n_vars

    for v_idx in range(n_vars):
        st.session_state[f"run_{run_idx}_v_qty_{v_idx}"] = base_share + (1 if v_idx < remainder else 0)


def on_dieline_or_variant_count_change(run_idx: int):
    """Triggered when a dieline changes or the number of variants within that dieline changes."""
    distribute_moq_evenly(run_idx)


def reset_all_form_fields():
    """Purges session state to reset the entire quote cleanly."""
    prefixes = ("run_die_", "run_n_vars_", "run_")
    known_keys = {
        "num_dieline_runs", "substrate_colour", "substrate_type", "material_specs",
        "colour_match_type", "qty_pantone_match", "pantone_refs",
        "qty_target_match", "target_refs", "litho_print", "digital_print",
        "epson_proof_qty", "layflat_supplied", "cm_net", "cm_qty",
        "lf_net", "lf_qty", "pt_net", "pt_qty", "preset_die_selector"
    }
    for key in list(st.session_state.keys()):
        if key in known_keys or key.startswith(prefixes):
            del st.session_state[key]


def load_single_quote_preset():
    """Loads a selected Ren-Flex digital trial quote as Run #1."""
    selection = st.session_state.get("preset_die_selector")
    if not selection or selection == "-- Select a Ren-Flex Trial Quote --":
        return

    reset_all_form_fields()
    st.session_state["num_dieline_runs"] = 1
    st.session_state["run_die_0"] = selection
    st.session_state["run_n_vars_0"] = 1
    distribute_moq_evenly(0)


# -----------------------------------------------------------------------------
# STREAMLIT UI
# -----------------------------------------------------------------------------

st.set_page_config(
    page_title="IML Digital Trial Quotation Form", page_icon="🏷️", layout="wide"
)

st.title("🏷️ In-Mould Label (IML) Digital Trial Quotation")
st.markdown(
    "Each digital trial quote is charged as **one MOQ press run per dieline size**, which can be split across **multiple label designs/variants** (in equal or custom quantities) as long as the variant quantities sum to the dieline's **MOQ**."
)

# Master Reference Table
with st.expander("📋 View Ren-Flex Digital Trial MOQ & Pricing Schedule (Quotes 10794–10801)", expanded=False):
    ref_df = pd.DataFrame([
        {
            "Quote Ref": v["quote_ref"],
            "Die": v["die_code"],
            "Jar / Format": v["jar_name"],
            "Label Dimensions": v["size_mm"],
            "Trial MOQ (Total Run)": f"{v['moq']:,}",
            "Unit Price (Excl. VAT)": f"R {v['unit_price']:.4f}",
            "Single Trial Run Charge": f"R {v['run_subtotal']:,.2f}",
        }
        for v in DIELINE_SPECS.values()
    ])
    st.dataframe(ref_df, use_container_width=True, hide_index=True)

# Top Controls
ctrl_col1, ctrl_col2 = st.columns([3, 1])
with ctrl_col1:
    st.selectbox(
        "📄 Quick-Load Single Ren-Flex Trial Quote:",
        options=["-- Select a Ren-Flex Trial Quote --"] + DIELINE_OPTIONS,
        key="preset_die_selector",
        on_change=load_single_quote_preset,
    )
with ctrl_col2:
    st.write("")
    st.write("")
    st.button("🔄 Refresh / Clear Form", on_click=reset_all_form_fields, use_container_width=True)

st.divider()

# -----------------------------------------------------------------------------
# 1. DIGITAL TRIAL DIELINE RUNS & VARIANT ALLOCATION
# -----------------------------------------------------------------------------
st.header("1. Digital Trial Dieline Runs & Variant Split")

num_dieline_runs = st.number_input(
    "Number of Dieline Sizes to Trial",
    min_value=1,
    max_value=8,
    step=1,
    value=1,
    key="num_dieline_runs",
    help="Each dieline has its own MOQ and flat trial run price. Increase this only if trialling multiple jar sizes."
)

dieline_runs_data = []
all_runs_valid = True
total_order_qty = 0
total_variant_count = 0
total_trial_runs_charge = 0.0
selected_dieline_names = []

for r in range(int(num_dieline_runs)):
    if f"run_die_{r}" not in st.session_state:
        st.session_state[f"run_die_{r}"] = DIELINE_OPTIONS[r % len(DIELINE_OPTIONS)]
    if f"run_n_vars_{r}" not in st.session_state:
        st.session_state[f"run_n_vars_{r}"] = 1

    current_die = st.session_state[f"run_die_{r}"]
    spec = DIELINE_SPECS[current_die]
    required_moq = spec["moq"]
    unit_price = spec["unit_price"]
    run_charge = required_moq * unit_price

    with st.container(border=True):
        st.subheader(f"Dieline Trial Run #{r+1}")

        top_c1, top_c2, top_c3, top_c4 = st.columns([3, 1.3, 1.3, 1.4])
        with top_c1:
            chosen_die = st.selectbox(
                "Select Dieline / Jar Size",
                options=DIELINE_OPTIONS,
                key=f"run_die_{r}",
                on_change=on_dieline_or_variant_count_change,
                args=(r,),
            )
            st.caption(f"Quote Ref: **{spec['quote_ref']}** | Size: **{spec['size_mm']}**")

        with top_c2:
            n_vars = st.number_input(
                "No. of Variants in Die",
                min_value=1,
                max_value=20,
                step=1,
                key=f"run_n_vars_{r}",
                on_change=on_dieline_or_variant_count_change,
                args=(r,),
            )

        with top_c3:
            st.metric("Required MOQ (Run Qty)", f"{required_moq:,}")

        with top_c4:
            st.metric(
                "Single Run Charge",
                f"R {run_charge:,.2f}",
                delta=f"@ R {unit_price:.4f} / label",
                delta_color="off"
            )

        st.markdown("**Specify Label Designs & Quantities within this Die Size:**")

        # Ensure variant quantities exist for this run
        if f"run_{r}_v_qty_0" not in st.session_state:
            distribute_moq_evenly(r)

        run_variants = []
        allocated_sum = 0

        for v in range(int(n_vars)):
            if f"run_{r}_v_qty_{v}" not in st.session_state:
                st.session_state[f"run_{r}_v_qty_{v}"] = 0

            vc1, vc2, vc3 = st.columns([3.5, 1.5, 1])
            with vc1:
                v_desc = st.text_input(
                    f"Design {v+1} Description / Artwork Name",
                    key=f"run_{r}_v_name_{v}",
                    placeholder=f'e.g., "{spec["jar_name"]} - Flavour / Variant {v+1}"',
                )
            with vc2:
                v_qty = st.number_input(
                    f"Design {v+1} Qty",
                    min_value=0,
                    max_value=required_moq,
                    step=500,
                    key=f"run_{r}_v_qty_{v}",
                )
            with vc3:
                pct_share = (v_qty / required_moq * 100) if required_moq > 0 else 0
                st.write("**Share**")
                st.write(f"{pct_share:.1f}%")

            allocated_sum += v_qty
            run_variants.append({
                "variant": v_desc,
                "quantity": v_qty,
            })

        # MOQ Validation & Auto-Balance Bar
        diff = required_moq - allocated_sum
        val_c1, val_c2 = st.columns([3.5, 1.5])
        with val_c1:
            if diff == 0:
                st.success(f"✅ **Allocated: {allocated_sum:,} / {required_moq:,} units** — Exact match to Dieline MOQ.")
            elif diff > 0:
                all_runs_valid = False
                st.error(f"⚠️ **Allocated: {allocated_sum:,} / {required_moq:,} units** — Please allocate **{diff:,} more units** to reach the MOQ.")
            else:
                all_runs_valid = False
                st.error(f"⚠️ **Allocated: {allocated_sum:,} / {required_moq:,} units** — Over MOQ by **{abs(diff):,} units**. Reduce variant quantities.")
        with val_c2:
            st.button(
                "⚖️ Split MOQ Evenly",
                key=f"balance_btn_{r}",
                on_click=distribute_moq_evenly,
                args=(r,),
                use_container_width=True,
            )

    total_order_qty += required_moq
    total_variant_count += int(n_vars)
    total_trial_runs_charge += run_charge
    if chosen_die not in selected_dieline_names:
        selected_dieline_names.append(chosen_die)

    dieline_runs_data.append({
        "dieline": chosen_die,
        "quote_ref": spec["quote_ref"],
        "die_code": spec["die_code"],
        "jar_name": spec["jar_name"],
        "size_mm": spec["size_mm"],
        "moq": required_moq,
        "unit_price": unit_price,
        "run_charge": run_charge,
        "variants": run_variants,
    })

st.divider()

# -----------------------------------------------------------------------------
# 2. SUBSTRATE SPECIFICATIONS
# -----------------------------------------------------------------------------
st.header("2. Substrate Specifications")

col_sub1, col_sub2 = st.columns(2)
with col_sub1:
    substrate_colour = st.radio(
        "Substrate Colour",
        options=["White", "Clear"],
        horizontal=True,
        key="substrate_colour"
    )
    substrate_type = st.radio(
        "Substrate Type",
        options=["Smooth", "Orange Peel"],
        horizontal=True,
        key="substrate_type"
    )

with col_sub2:
    if "material_specs" not in st.session_state:
        st.session_state["material_specs"] = DEFAULT_MATERIAL_SPECS
    material_specs = st.text_area(
        "Material Specifications",
        height=120,
        key="material_specs"
    )

st.divider()

# -----------------------------------------------------------------------------
# 3. JAR COMPATIBILITY (SYNCED WITH TRIAL DIELINES)
# -----------------------------------------------------------------------------
st.header("3. IML Labels Required for Jars")
st.caption("Pre-selected from the dielines configured in Section 1:")

selected_jars = st.multiselect(
    "Target Jars & Dieline Dimensions",
    options=DIELINE_OPTIONS,
    default=selected_dieline_names,
)

st.divider()

# -----------------------------------------------------------------------------
# 4. COLOUR MATCHING & PRINT METHOD
# -----------------------------------------------------------------------------
st.header("4. Colour Matching & Print Method")

col_c1, col_c2 = st.columns(2)
with col_c1:
    colour_match_type = st.multiselect(
        "Colour Matching Method Required",
        options=["Pantone", "Match to target supplied"],
        key="colour_match_type"
    )

    qty_pantone_match = 0
    pantone_refs = ""
    if "Pantone" in colour_match_type:
        qty_pantone_match = st.number_input(
            "No. of labels to match to Pantone (Qty)", min_value=0, step=1, key="qty_pantone_match"
        )
        pantone_refs = st.text_input(
            "Specify Pantone References",
            placeholder='e.g., "PMS 185 C, PMS 286 C"',
            key="pantone_refs"
        )

    qty_target_match = 0
    target_refs = ""
    if "Match to target supplied" in colour_match_type:
        qty_target_match = st.number_input(
            "No. of labels to match to Target (Qty)", min_value=0, step=1, key="qty_target_match"
        )
        target_refs = st.text_input(
            "Specify Target References / Samples",
            placeholder='e.g., "Physical sample ref #1042"',
            key="target_refs"
        )

with col_c2:
    st.write("**Print Process & Proofing**")
    litho_print = st.radio(
        "Litho Print Required?", options=["No", "Yes"], horizontal=True, key="litho_print"
    )
    digital_print = st.radio(
        "Digital Print Required?", options=["No", "Yes"], horizontal=True, index=1, key="digital_print"
    )

    epson_proof_qty = st.number_input(
        "Epson Proof (Qty)", min_value=0, step=1, value=1, key="epson_proof_qty"
    )
    layflat_supplied = st.radio(
        "Layflat for layout and colour to be supplied?",
        options=["No", "Yes"],
        horizontal=True,
        key="layflat_supplied"
    )

st.divider()

# -----------------------------------------------------------------------------
# 5. FINANCIAL QUOTATION SUMMARY
# -----------------------------------------------------------------------------
st.header("5. Financial Quotation Summary")
st.caption("Each Dieline Trial Run is charged as a single MOQ run. Add optional pre-press/proofing charges below:")

col_desc, col_val, col_qty, col_sub = st.columns([3.5, 2, 1.5, 2])
col_desc.write("**Task Description**")
col_val.write("**Unit Price**")
col_qty.write("**MOQ Qty**")
col_sub.write("**Sub Total**")

# Auto-populated Digital Trial Runs (1 line per dieline run)
for run in dieline_runs_data:
    r_desc, r_val, r_qty, r_sub = st.columns([3.5, 2, 1.5, 2])
    r_desc.write(f"Digital Trial Run: **{run['die_code']} — {run['jar_name']}** ({len(run['variants'])} variant(s))")
    r_val.write(f"R {run['unit_price']:.4f}")
    r_qty.write(f"{run['moq']:,}")
    r_sub.markdown(f"**R {run['run_charge']:,.2f}**")

# Optional Extra Charges
c_desc, c_val, c_qty, c_sub = st.columns([3.5, 2, 1.5, 2])
c_desc.write("Colour Matching")
cm_net_val = c_val.number_input("CM Net Value", min_value=0.0, value=0.0, step=10.0, key="cm_net", label_visibility="collapsed")
cm_quantity = c_qty.number_input("CM Quantity", min_value=0, value=0, step=1, key="cm_qty", label_visibility="collapsed")
cm_subtotal = cm_net_val * cm_quantity
c_sub.markdown(f"**R {cm_subtotal:,.2f}**")

l_desc, l_val, l_qty, l_sub = st.columns([3.5, 2, 1.5, 2])
l_desc.write("Layflat")
lf_net_val = l_val.number_input("LF Net Value", min_value=0.0, value=0.0, step=10.0, key="lf_net", label_visibility="collapsed")
lf_quantity = l_qty.number_input("LF Quantity", min_value=0, value=0, step=1, key="lf_qty", label_visibility="collapsed")
lf_subtotal = lf_net_val * lf_quantity
l_sub.markdown(f"**R {lf_subtotal:,.2f}**")

p_desc, p_val, p_qty, p_sub = st.columns([3.5, 2, 1.5, 2])
p_desc.write("Additional Setup / Tooling")
pt_net_val = p_val.number_input("PT Net Value", min_value=0.0, value=0.0, step=10.0, key="pt_net", label_visibility="collapsed")
pt_quantity = p_qty.number_input("PT Quantity", min_value=0, value=0, step=1, key="pt_qty", label_visibility="collapsed")
pt_subtotal = pt_net_val * pt_quantity
p_sub.markdown(f"**R {pt_subtotal:,.2f}**")

st.divider()

# Live Totals Calculation
net_subtotal = total_trial_runs_charge + cm_subtotal + lf_subtotal + pt_subtotal
vat_amount = net_subtotal * 0.15
gross_total = net_subtotal + vat_amount

tot_col1, tot_col2 = st.columns([6, 3])
with tot_col2:
    st.write(f"**Net Subtotal:** R {net_subtotal:,.2f}")
    st.write(f"**15% VAT:** R {vat_amount:,.2f}")
    st.subheader(f"Gross Total: R {gross_total:,.2f}")

st.divider()

# SUBMIT BUTTON (Disabled if any dieline run's variants do not sum to its MOQ)
submit_button = st.button(
    "Submit Digital Trial Quotation & Generate PDF",
    type="primary",
    use_container_width=True,
    disabled=not all_runs_valid,
)

if not all_runs_valid:
    st.warning("⚠️ Please ensure the variant quantities in each Dieline Trial Run equal the required MOQ before submitting.")

# -----------------------------------------------------------------------------
# POST-SUBMISSION ACTIONS
# -----------------------------------------------------------------------------
if submit_button and all_runs_valid:
    quote_data = {
        "Total Quantity": total_order_qty,
        "Total Variants": total_variant_count,
        "Dieline Runs": dieline_runs_data,
        "Substrate Colour": substrate_colour,
        "Substrate Type": substrate_type,
        "Material Specs": material_specs,
        "Target Jars": selected_jars,
        "Colour Matching": colour_match_type,
        "Pantone Qty": qty_pantone_match,
        "Pantone Refs": pantone_refs,
        "Target Qty": qty_target_match,
        "Target Refs": target_refs,
        "Litho Print": litho_print,
        "Digital Print": digital_print,
        "Epson Proof Qty": epson_proof_qty,
        "Layflat Supplied": layflat_supplied,
        "cm_net_val": cm_net_val,
        "cm_quantity": cm_quantity,
        "lf_net_val": lf_net_val,
        "lf_quantity": lf_quantity,
        "pt_net_val": pt_net_val,
        "pt_quantity": pt_quantity,
        "Net Subtotal": net_subtotal,
        "VAT (15%)": vat_amount,
        "Gross Total": gross_total,
    }

    st.success("🎉 Digital Trial Quotation processed successfully!")

    # 1. Update Google Spreadsheet
    with st.spinner("Recording request in Google Spreadsheet..."):
        sheet_updated = append_to_google_sheets(quote_data)
        if sheet_updated:
            st.info("✅ Data appended to Google Sheet database.")

    # 2. Generate PDF Document & Display Download Button
    with st.spinner("Generating client PDF quotation document..."):
        pdf_bytes = generate_pdf_quote(quote_data)
        
        if pdf_bytes and HAS_FPDF:
            filename = f"IML_Digital_Trial_Quote_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
            
            st.download_button(
                label="📥 Download Quotation PDF",
                data=pdf_bytes,
                file_name=filename,
                mime="application/pdf",
                use_container_width=True
            )
        elif not HAS_FPDF:
            st.warning("`fpdf2` is not installed. Install it using `pip install fpdf2` to enable PDF downloads.")