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

        row = [
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            data["Total Quantity"],
            data["No of Variants"],
            str([f"{v['variant']}: {v['quantity']}" for v in data["Variant Details"]]),
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
            data["Trials Required"],
            data["Trial Variants"],
            data["Trial Qty Per Variant"],
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
    """
    Generates a clean PDF document for the client using FPDF.
    """
    if not HAS_FPDF:
        return b"PDF Generation requires 'fpdf2' package."

    class PDF(FPDF):
        def header(self):
            self.set_font("Helvetica", "B", 16)
            self.set_text_color(30, 41, 59)
            self.cell(0, 10, "In-Mould Label (IML) Quotation", border=False, new_x="LMARGIN", new_y="NEXT", align="L")
            self.set_font("Helvetica", "", 9)
            self.set_text_color(100, 116, 139)
            self.cell(0, 5, f"Generated on: {datetime.now().strftime('%d %B %Y')} | Quote Ref: #IML-{datetime.now().strftime('%M%S')}", border=False, new_x="LMARGIN", new_y="NEXT", align="L")
            self.ln(5)

        def footer(self):
            self.set_y(-15)
            self.set_font("Helvetica", "I", 8)
            self.set_text_color(148, 163, 184)
            self.cell(0, 10, f"Page {self.page_no()} | Confidential IML Quote", align="C")

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
        pdf.cell(0, 6, str(value), border=False, new_x="LMARGIN", new_y="NEXT")

    # 1. Summary Overview
    section_heading("1. Quantity & Range Overview")
    key_value("Total Range Quantity:", f"{data['Total Quantity']:,} units")
    key_value("Number of Variants:", data["No of Variants"])
    pdf.ln(2)

    # Variant Table
    pdf.set_font("Helvetica", "B", 8)
    pdf.set_fill_color(226, 232, 240)
    pdf.cell(120, 6, "Variant Name / Description", border=1, fill=True)
    pdf.cell(60, 6, "Quantity", border=1, new_x="LMARGIN", new_y="NEXT", fill=True)
    
    pdf.set_font("Helvetica", "", 8)
    for v in data["Variant Details"]:
        v_name = v["variant"] if v["variant"] else "Unnamed Variant"
        pdf.cell(120, 6, v_name, border=1)
        pdf.cell(60, 6, f"{v['quantity']:,}", border=1, new_x="LMARGIN", new_y="NEXT")
    pdf.ln(4)

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

    # 5. Production Trials
    section_heading("5. Production Trials")
    key_value("Trials Required:", data["Trials Required"])
    if data["Trials Required"] == "Yes":
        key_value("Trial Variants:", data["Trial Variants"])
        key_value("Trial Qty per Variant:", f"{data['Trial Qty Per Variant']:,} units")
    pdf.ln(4)

    # 6. Financial Summary Table
    section_heading("6. Financial Summary")
    
    pdf.set_font("Helvetica", "B", 8)
    pdf.set_fill_color(226, 232, 240)
    pdf.cell(80, 6, "Task Description", border=1, fill=True)
    pdf.cell(35, 6, "Net Value", border=1, align="R", fill=True)
    pdf.cell(30, 6, "Qty", border=1, align="R", fill=True)
    pdf.cell(35, 6, "Sub Total", border=1, new_x="LMARGIN", new_y="NEXT", align="R", fill=True)

    pdf.set_font("Helvetica", "", 8)
    financial_items = [
        ("Colour Matching", data.get("cm_net_val", 0.0), data.get("cm_quantity", 0)),
        ("Layflat", data.get("lf_net_val", 0.0), data.get("lf_quantity", 0)),
        ("Production Trials", data.get("pt_net_val", 0.0), data.get("pt_quantity", 0)),
    ]

    for item_desc, net_val, qty in financial_items:
        subtotal = net_val * qty
        pdf.cell(80, 6, item_desc, border=1)
        pdf.cell(35, 6, f"R {net_val:,.2f}", border=1, align="R")
        pdf.cell(30, 6, str(qty), border=1, align="R")
        pdf.cell(35, 6, f"R {subtotal:,.2f}", border=1, new_x="LMARGIN", new_y="NEXT", align="R")

    pdf.ln(2)
    pdf.set_font("Helvetica", "B", 9)
    pdf.cell(145, 6, "Net Subtotal:", border=0, align="R")
    pdf.cell(35, 6, f"R {data['Net Subtotal']:,.2f}", border=1, new_x="LMARGIN", new_y="NEXT", align="R")

    pdf.cell(145, 6, "15% VAT:", border=0, align="R")
    pdf.cell(35, 6, f"R {data['VAT (15%)']:,.2f}", border=1, new_x="LMARGIN", new_y="NEXT", align="R")

    pdf.set_fill_color(241, 245, 249)
    pdf.cell(145, 7, "Gross Total:", border=0, align="R")
    pdf.cell(35, 7, f"R {data['Gross Total']:,.2f}", border=1, new_x="LMARGIN", new_y="NEXT", align="R", fill=True)

    # Expressly cast bytearray output to standard bytes for Streamlit download_button
    return bytes(pdf.output())


# -----------------------------------------------------------------------------
# FORM RESET CALLBACK FUNCTION
# -----------------------------------------------------------------------------

def reset_all_form_fields():
    """Callback function triggered by the Refresh button to purge session state."""
    known_keys = [
        "total_qty", "num_variants", "substrate_colour", "substrate_type", 
        "material_specs", "selected_jars", "colour_match_type", "qty_pantone_match", 
        "pantone_refs", "qty_target_match", "target_refs", "litho_print", 
        "digital_print", "epson_proof_qty", "layflat_supplied", "trials_required", 
        "trial_variants", "trial_qty_per_variant", "cm_net", "cm_qty", 
        "lf_net", "lf_qty", "pt_net", "pt_qty"
    ]
    
    # Catch dynamic variant input keys
    for key in list(st.session_state.keys()):
        if key.startswith("v_name_") or key.startswith("v_qty_"):
            known_keys.append(key)
            
    for key in known_keys:
        if key in st.session_state:
            del st.session_state[key]


# -----------------------------------------------------------------------------
# STREAMLIT UI & FORM
# -----------------------------------------------------------------------------

st.set_page_config(
    page_title="IML Label Quotation Form", page_icon="🏷️", layout="wide"
)

st.title("🏷️ In-Mould Label (IML) Quotation Request")
st.markdown(
    "Please fill out the details below to generate an accurate quotation for your IML label range."
)

# Refresh Button with on_click Callback
st.button("🔄 Refresh / Clear Form", on_click=reset_all_form_fields, use_container_width=True)

with st.form("iml_quote_form"):

    # 1. QUANTITIES & VARIANTS
    st.header("1. Quantities & Variant Breakdown")

    col1, col2 = st.columns(2)
    with col1:
        total_qty = st.number_input(
            "Total Qty of IML Labels in Range", min_value=1, step=100, value=10000, key="total_qty"
        )
    with col2:
        num_variants = st.number_input(
            "No. of IML Label Variants in Range", min_value=1, step=1, value=1, key="num_variants"
        )

    st.subheader("Variant Quantities & Descriptions")
    st.info("Specify the details and quantities for each individual variant in the range.")

    variant_details = []
    for i in range(int(num_variants)):
        v_col1, v_col2 = st.columns([2, 1])
        with v_col1:
            v_name = st.text_input(
                f"Variant {i+1} Name / Description",
                key=f"v_name_{i}",
                placeholder='e.g., "Vanilla 250g"',
            )
        with v_col2:
            v_qty = st.number_input(
                f"Variant {i+1} Qty", min_value=1, step=500, key=f"v_qty_{i}"
            )
        variant_details.append({"variant": v_name, "quantity": v_qty})

    st.divider()

    # 2. SUBSTRATE SPECIFICATIONS
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
        material_specs = st.text_area(
            "Material Specifications",
            placeholder="Specify film thickness, barrier properties, or specific resin grade requirements...",
            height=120,
            key="material_specs"
        )

    st.divider()

    # 3. JAR COMPATIBILITY
    st.header("3. IML Labels Required for Jars")
    st.caption("Select all jar formats required for this quotation:")

    jar_options = [
        "150mm non-tamper lock jar",
        "250mm non-tamper lock jar",
        "350mm non-tamper lock jar",
        "450mm non-tamper lock jar",
        "450mm standard lightweight jar",
        "250mm Eve jar",
        "350mm Eve jar",
        "500mm jar with standard length label",
        "500mm jar with standard length label plus 1mm longer label (Clicks / custom specification)",
    ]

    selected_jars = st.multiselect(
        "Select Target Jars (Specify)", options=jar_options, key="selected_jars"
    )

    st.divider()

    # 4. COLOUR MATCHING & PROCESS
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
            "Digital Print Required?", options=["No", "Yes"], horizontal=True, key="digital_print"
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

    # 5. TRIALS
    st.header("5. Production Trials")

    trials_required = st.radio(
        "Trials Required?", options=["No", "Yes"], horizontal=True, key="trials_required"
    )

    trial_variants = 0
    trial_qty_per_variant = 0

    if trials_required == "Yes":
        col_t1, col_t2 = st.columns(2)
        with col_t1:
            trial_variants = st.number_input(
                "Trials (No. of variants)", min_value=1, step=1, value=1, key="trial_variants"
            )
        with col_t2:
            trial_qty_per_variant = st.number_input(
                "Trials (Quantity per variant)",
                min_value=1,
                step=100,
                value=500,
                key="trial_qty_per_variant"
            )

    st.divider()

    # 6. FINANCIAL COST SUMMARY
    st.header("6. Financial Quotation Summary")
    st.caption("Specify net values and quantities for task items:")

    col_desc, col_val, col_qty, col_sub = st.columns([3, 2, 2, 2])
    col_desc.write("**Task Description**")
    col_val.write("**Net Value per Unit**")
    col_qty.write("**Quantity**")
    col_sub.write("**Sub Total**")

    # Item 1: Colour Matching
    c_desc, c_val, c_qty, c_sub = st.columns([3, 2, 2, 2])
    c_desc.write("Colour Matching")
    cm_net_val = c_val.number_input("Net Value", min_value=0.0, value=0.0, step=10.0, key="cm_net", label_visibility="collapsed")
    cm_quantity = c_qty.number_input("Quantity", min_value=0, value=0, step=1, key="cm_qty", label_visibility="collapsed")
    cm_subtotal = cm_net_val * cm_quantity
    c_sub.markdown(f"**R {cm_subtotal:,.2f}**")

    # Item 2: Layflat
    l_desc, l_val, l_qty, l_sub = st.columns([3, 2, 2, 2])
    l_desc.write("Layflat")
    lf_net_val = l_val.number_input("Net Value", min_value=0.0, value=0.0, step=10.0, key="lf_net", label_visibility="collapsed")
    lf_quantity = l_qty.number_input("Quantity", min_value=0, value=0, step=1, key="lf_qty", label_visibility="collapsed")
    lf_subtotal = lf_net_val * lf_quantity
    l_sub.markdown(f"**R {lf_subtotal:,.2f}**")

    # Item 3: Production Trials
    p_desc, p_val, p_qty, p_sub = st.columns([3, 2, 2, 2])
    p_desc.write("Production Trials")
    pt_net_val = p_val.number_input("Net Value", min_value=0.0, value=0.0, step=10.0, key="pt_net", label_visibility="collapsed")
    pt_quantity = p_qty.number_input("Quantity", min_value=0, value=0, step=1, key="pt_qty", label_visibility="collapsed")
    pt_subtotal = pt_net_val * pt_quantity
    p_sub.markdown(f"**R {pt_subtotal:,.2f}**")

    st.divider()

    # Totals Calculation
    net_subtotal = cm_subtotal + lf_subtotal + pt_subtotal
    vat_amount = net_subtotal * 0.15
    gross_total = net_subtotal + vat_amount

    tot_col1, tot_col2 = st.columns([6, 3])
    with tot_col2:
        st.write(f"**Net Subtotal:** R {net_subtotal:,.2f}")
        st.write(f"**15% VAT:** R {vat_amount:,.2f}")
        st.subheader(f"Gross Total: R {gross_total:,.2f}")

    st.divider()

    # SUBMIT BUTTON
    submit_button = st.form_submit_button(
        "Submit Quotation Request", use_container_width=True
    )

# -----------------------------------------------------------------------------
# POST-SUBMISSION ACTIONS
# -----------------------------------------------------------------------------
if submit_button:
    quote_data = {
        "Total Quantity": total_qty,
        "No of Variants": num_variants,
        "Variant Details": variant_details,
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
        "Trials Required": trials_required,
        "Trial Variants": trial_variants if trials_required == "Yes" else 0,
        "Trial Qty Per Variant": (
            trial_qty_per_variant if trials_required == "Yes" else 0
        ),
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

    st.success("🎉 Quotation request processed successfully!")

    # 1. Update Google Spreadsheet
    with st.spinner("Recording request in Google Spreadsheet..."):
        sheet_updated = append_to_google_sheets(quote_data)
        if sheet_updated:
            st.info("✅ Data appended to Google Sheet database.")

    # 2. Generate PDF Document & Display Download Button
    with st.spinner("Generating client PDF quotation document..."):
        pdf_bytes = generate_pdf_quote(quote_data)
        
        if pdf_bytes and HAS_FPDF:
            filename = f"IML_Quotation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
            
            st.download_button(
                label="📥 Download Quotation PDF",
                data=pdf_bytes,
                file_name=filename,
                mime="application/pdf",
                use_container_width=True
            )
        elif not HAS_FPDF:
            st.warning("`fpdf2` is not installed. Install it using `pip install fpdf2` to enable PDF downloads.")