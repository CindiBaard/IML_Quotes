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
# DIGITAL TRIAL DIELINE MASTER DATA & PRICING CONSTANTS
# -----------------------------------------------------------------------------

CLEAR_MARKUP_MULTIPLIER = 1.05  # +5% substrate markup for Clear IML
EPSON_PROOF_UNIT_COST = 650.00  # R 650.00 base cost per Epson proof

RENFLEX_SUPPLIER_DETAILS = {
    "name": "REN-MEDIA CC t/a REN-FLEX",
    "reg_no": "2002/030011/23",
    "vat_no": "4900228500",
    "address": "37 Kimberley Booysens, Cnr Garland Road, Ophirton, Johannesburg, 2001 / PO Box 1125, Southdale 2135",
    "tel": "(011) 493 1575",
    "rep": "Francois",
}

DIELINE_SPECS = {
    "Die 173 - 250ml Jar (274.395mm x 44.121mm)": {
        "quote_ref": "10794/01",
        "die_code": "Die 173",
        "jar_name": "250ml Jar",
        "size_mm": "274.395mm x 44.121mm",
        "moq": 25000,
        "base_unit_price": 0.5162,
    },
    "Die 174 - 350ml Jar (274.392mm x 60.183mm)": {
        "quote_ref": "10795/01",
        "die_code": "Die 174",
        "jar_name": "350ml Jar",
        "size_mm": "274.392mm x 60.183mm",
        "moq": 18000,
        "base_unit_price": 0.7082,
    },
    "Die 175 - 450ml Jar (274.387mm x 82.041mm)": {
        "quote_ref": "10796/01",
        "die_code": "Die 175",
        "jar_name": "450ml Jar",
        "size_mm": "274.387mm x 82.041mm",
        "moq": 13500,
        "base_unit_price": 0.9650,
    },
    "Die 176 - 500ml Jar (291.273mm x 84.605mm)": {
        "quote_ref": "10797/01",
        "die_code": "Die 176",
        "jar_name": "500ml Jar",
        "size_mm": "291.273mm x 84.605mm",
        "moq": 13500,
        "base_unit_price": 0.9650,
    },
    "Die 189 - 450ml Std Lightweight Jar (275.378mm x 87.066mm)": {
        "quote_ref": "10798/01",
        "die_code": "Die 189",
        "jar_name": "450ml Std Lightweight Jar",
        "size_mm": "275.378mm x 87.066mm",
        "moq": 13500,
        "base_unit_price": 0.9650,
    },
    "Die 211 - 500ml Eve Clicks Tub (292.304mm x 81.936mm)": {
        "quote_ref": "10799/01",
        "die_code": "Die 211",
        "jar_name": "500ml Eve Clicks Tub",
        "size_mm": "292.304mm x 81.936mm",
        "moq": 13500,
        "base_unit_price": 0.9650,
    },
    "Die 216 - Clicks Labels (292.304mm x 81.936mm)": {
        "quote_ref": "10800/01",
        "die_code": "Die 216",
        "jar_name": "Clicks Labels",
        "size_mm": "292.304mm x 81.936mm",
        "moq": 13500,
        "base_unit_price": 0.9650,
    },
    "Die 264 - 250ml Fabupharm (269.908mm x 44.967mm)": {
        "quote_ref": "10801/01",
        "die_code": "Die 264",
        "jar_name": "250ml Fabupharm",
        "size_mm": "269.908mm x 44.967mm",
        "moq": 25000,
        "base_unit_price": 0.5162,
    },
}

DIELINE_OPTIONS = list(DIELINE_SPECS.keys())

WHITE_MATERIAL_SPECS = (
    "Printed CMYK front only on IML White Gloss, 60mic | "
    "Finishing: Waterbased Varnish, Die Cutting - Digital | Other Material: Primer | "
    "Note: Subject to +/-10% over/under delivery"
)

CLEAR_MATERIAL_SPECS = (
    "Printed CMYK + double -white front only on IML Clear, 60mic | "
    "Finishing: Waterbased Varnish, Die Cutting - Digital | Other Material: Primer | "
    "Note: Subject to +/-10% over/under delivery"
)


def get_effective_unit_price(base_price: float, substrate_colour: str) -> float:
    """Returns the supplier unit price with a +5% markup if Clear substrate is selected."""
    if substrate_colour == "Clear":
        return round(base_price * CLEAR_MARKUP_MULTIPLIER, 4)
    return base_price


def pdf_safe_text(text) -> str:
    """Sanitizes characters (such as em-dashes) so core Helvetica FPDF fonts never fail."""
    return str(text).replace("—", "-").replace("–", "-")


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

        variant_list = [f"{v['variant']}: {v['quantity']}" for v in data["Variant Details"]]
        doc_tag = f"ORDER ({data['PO Number']})" if data["Document Type"] == "Ren-Flex Order" else "CLIENT QUOTE"

        row = [
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            data["Total Quantity"],
            data["No of Variants"],
            str(variant_list),
            data["Substrate Colour"],
            data["Substrate Type"],
            data["Material Specs"],
            data["Selected Dieline"],
            ", ".join(data["Colour Matching"]),
            data["Pantone Qty"],
            data["Pantone Refs"],
            data["Target Qty"],
            data["Target Refs"],
            data["Litho Print"],
            data["Digital Print"],
            data["Epson Proof Qty"],
            data["Layflat Supplied"],
            f"Yes ({doc_tag} | Base Net: R{data['Base Net Subtotal']:,.2f})",
            data["No of Variants"],
            "Split across variants to MOQ",
            data["Client Net Subtotal"],
            data["Client VAT (15%)"],
            data["Client Gross Total"]
        ]

        sheet.append_row(row)
        return True
    except Exception as e:
        st.error(f"Failed to append to Google Sheet: {e}")
        return False


def generate_client_pdf_quote(data: dict) -> bytes:
    """Generates the Client Quotation PDF using the marked-up client selling prices."""
    if not HAS_FPDF:
        return b"PDF Generation requires 'fpdf2' package."

    class PDF(FPDF):
        def header(self):
            self.set_font("Helvetica", "B", 16)
            self.set_text_color(30, 41, 59)
            self.cell(0, 10, "In-Mould Label (IML) Digital Trial Quotation", border=False, new_x="LMARGIN", new_y="NEXT", align="L")
            self.set_font("Helvetica", "", 9)
            self.set_text_color(100, 116, 139)
            client_str = f" | Client: {pdf_safe_text(data['Client Name'])}" if data.get("Client Name") else ""
            self.cell(0, 5, f"Generated on: {datetime.now().strftime('%d %B %Y')} | Quote Ref: #IML-Q-{datetime.now().strftime('%M%S')}{client_str}", border=False, new_x="LMARGIN", new_y="NEXT", align="L")
            self.ln(5)

        def footer(self):
            self.set_y(-15)
            self.set_font("Helvetica", "I", 8)
            self.set_text_color(148, 163, 184)
            self.cell(0, 10, f"Page {self.page_no()} | Valid for 30 days | Subject to +/-10% over/under delivery", align="C")

    pdf = PDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)

    def section_heading(title):
        pdf.set_x(pdf.l_margin)
        pdf.set_font("Helvetica", "B", 11)
        pdf.set_fill_color(241, 245, 249)
        pdf.set_text_color(15, 23, 42)
        pdf.cell(0, 7, f"  {pdf_safe_text(title)}", border=False, new_x="LMARGIN", new_y="NEXT", fill=True)
        pdf.ln(2)

    def key_value(key, value):
        pdf.set_x(pdf.l_margin)
        pdf.set_font("Helvetica", "B", 9)
        pdf.set_text_color(71, 85, 105)
        pdf.cell(60, 6, pdf_safe_text(key), border=False)
        pdf.set_font("Helvetica", "", 9)
        pdf.set_text_color(15, 23, 42)
        remaining_width = pdf.w - pdf.r_margin - pdf.get_x()
        pdf.multi_cell(remaining_width, 6, pdf_safe_text(value), border=False)

    # 1. Dieline & Variant Allocation
    section_heading("1. Dieline Specification & Variant Breakdown")
    if data.get("Client Name"):
        key_value("Prepared For:", data["Client Name"])
    key_value("Selected Dieline / Jar:", data["Selected Dieline"])
    key_value("Label Dimensions:", data["Dieline Size"])
    key_value("Total Trial MOQ:", f"{data['Total Quantity']:,} labels")
    key_value("Number of Variants:", data["No of Variants"])
    pdf.ln(2)

    pdf.set_x(pdf.l_margin)
    pdf.set_font("Helvetica", "B", 8)
    pdf.set_fill_color(226, 232, 240)
    pdf.cell(125, 6, "Label Design / Variant Description", border=1, fill=True)
    pdf.cell(35, 6, "Allocated Qty", border=1, align="R", fill=True)
    pdf.cell(30, 6, "% of MOQ", border=1, new_x="LMARGIN", new_y="NEXT", align="R", fill=True)

    pdf.set_font("Helvetica", "", 8)
    for v in data["Variant Details"]:
        pdf.set_x(pdf.l_margin)
        v_name = pdf_safe_text(v["variant"]) if v["variant"] else "Unnamed Label Design"
        pct = (v["quantity"] / data["Total Quantity"] * 100) if data["Total Quantity"] > 0 else 0
        pdf.cell(125, 6, v_name[:75], border=1)
        pdf.cell(35, 6, f"{v['quantity']:,}", border=1, align="R")
        pdf.cell(30, 6, f"{pct:.1f}%", border=1, new_x="LMARGIN", new_y="NEXT", align="R")
    pdf.ln(4)

    # 2. Substrate
    section_heading("2. Substrate Specifications")
    key_value("Substrate Colour:", data["Substrate Colour"])
    key_value("Substrate Finish:", data["Substrate Type"])
    key_value("Material Specs:", data["Material Specs"] if data["Material Specs"] else "Standard Specification")
    pdf.ln(4)

    # 3. Print & Colour
    section_heading("3. Print & Colour Specifications")
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

    # 4. Client Financial Summary Table (Marked-Up Client Prices)
    section_heading("4. Quotation Pricing Summary")
    
    pdf.set_x(pdf.l_margin)
    pdf.set_font("Helvetica", "B", 8)
    pdf.set_fill_color(226, 232, 240)
    pdf.cell(95, 6, "Description", border=1, fill=True)
    pdf.cell(35, 6, "Unit Price", border=1, align="R", fill=True)
    pdf.cell(25, 6, "Qty", border=1, align="R", fill=True)
    pdf.cell(35, 6, "Sub Total", border=1, new_x="LMARGIN", new_y="NEXT", align="R", fill=True)

    for item in data["Line Items"]:
        if item["qty"] > 0 or item["client_subtotal"] > 0:
            pdf.set_x(pdf.l_margin)
            pdf.set_font("Helvetica", "", 8)
            unit_fmt = f"R {item['client_unit_price']:,.4f}" if item["decimals"] == 4 else f"R {item['client_unit_price']:,.2f}"
            pdf.cell(95, 6, pdf_safe_text(item["description"])[:55], border=1)
            pdf.cell(35, 6, unit_fmt, border=1, align="R")
            pdf.cell(25, 6, f"{item['qty']:,}", border=1, align="R")
            pdf.cell(35, 6, f"R {item['client_subtotal']:,.2f}", border=1, new_x="LMARGIN", new_y="NEXT", align="R")

    pdf.ln(2)
    pdf.set_x(pdf.l_margin)
    pdf.set_font("Helvetica", "B", 9)
    pdf.cell(155, 6, "Net Subtotal:", border=0, align="R")
    pdf.cell(35, 6, f"R {data['Client Net Subtotal']:,.2f}", border=1, new_x="LMARGIN", new_y="NEXT", align="R")

    pdf.set_x(pdf.l_margin)
    pdf.cell(155, 6, "15% VAT:", border=0, align="R")
    pdf.cell(35, 6, f"R {data['Client VAT (15%)']:,.2f}", border=1, new_x="LMARGIN", new_y="NEXT", align="R")

    pdf.set_x(pdf.l_margin)
    pdf.set_fill_color(241, 245, 249)
    pdf.cell(155, 7, "Gross Total:", border=0, align="R")
    pdf.cell(35, 7, f"R {data['Client Gross Total']:,.2f}", border=1, new_x="LMARGIN", new_y="NEXT", align="R", fill=True)

    return bytes(pdf.output())


def generate_renflex_order_pdf(data: dict) -> bytes:
    """Generates a Supplier Purchase Order PDF for Ren-Flex Packaging at base supplier cost."""
    if not HAS_FPDF:
        return b"PDF Generation requires 'fpdf2' package."

    class PDF(FPDF):
        def header(self):
            self.set_font("Helvetica", "B", 16)
            self.set_text_color(15, 23, 42)
            self.cell(0, 10, "PURCHASE ORDER - IN-MOULD LABELS (DIGITAL TRIAL)", border=False, new_x="LMARGIN", new_y="NEXT", align="L")
            self.set_font("Helvetica", "", 9)
            self.set_text_color(71, 85, 105)
            self.cell(0, 5, f"Order Date: {datetime.now().strftime('%d %B %Y')} | PO No: {pdf_safe_text(data['PO Number'])} | Ren-Flex Quote Ref: {data['Quote Ref']}", border=False, new_x="LMARGIN", new_y="NEXT", align="L")
            self.ln(4)

        def footer(self):
            self.set_y(-15)
            self.set_font("Helvetica", "I", 8)
            self.set_text_color(148, 163, 184)
            self.cell(0, 10, f"Page {self.page_no()} | Purchase Order to REN-MEDIA CC t/a REN-FLEX", align="C")

    pdf = PDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)

    def section_heading(title):
        pdf.set_x(pdf.l_margin)
        pdf.set_font("Helvetica", "B", 10.5)
        pdf.set_fill_color(226, 232, 240)
        pdf.set_text_color(15, 23, 42)
        pdf.cell(0, 7, f"  {pdf_safe_text(title)}", border=False, new_x="LMARGIN", new_y="NEXT", fill=True)
        pdf.ln(2)

    def key_value(key, value):
        pdf.set_x(pdf.l_margin)
        pdf.set_font("Helvetica", "B", 9)
        pdf.set_text_color(71, 85, 105)
        pdf.cell(60, 5.5, pdf_safe_text(key), border=False)
        pdf.set_font("Helvetica", "", 9)
        pdf.set_text_color(15, 23, 42)
        remaining_width = pdf.w - pdf.r_margin - pdf.get_x()
        pdf.multi_cell(remaining_width, 5.5, pdf_safe_text(value), border=False)

    # 1. Supplier & Order Metadata
    section_heading("1. Supplier & Order Details")
    key_value("Supplier:", f"{RENFLEX_SUPPLIER_DETAILS['name']} (Reg: {RENFLEX_SUPPLIER_DETAILS['reg_no']})")
    key_value("Supplier VAT No:", RENFLEX_SUPPLIER_DETAILS["vat_no"])
    key_value("Attention / Rep:", RENFLEX_SUPPLIER_DETAILS["rep"])
    key_value("Ren-Flex Quote Ref:", data["Quote Ref"])
    key_value("Purchase Order No:", data["PO Number"])
    if data.get("Client Name"):
        key_value("End Customer / Ref:", data["Client Name"])
    if data.get("Delivery Address"):
        key_value("Delivery Instructions:", data["Delivery Address"])
    if data.get("Order Notes"):
        key_value("Artwork / Order Notes:", data["Order Notes"])
    pdf.ln(3)

    # 2. Dieline & Variant Split Instructions
    section_heading("2. Dieline & Variant Print Breakdown")
    key_value("Dieline / Tooling:", data["Selected Dieline"])
    key_value("Label Dimensions:", data["Dieline Size"])
    key_value("Total Order MOQ:", f"{data['Total Quantity']:,} labels across {data['No of Variants']} variant(s)")
    pdf.ln(2)

    pdf.set_x(pdf.l_margin)
    pdf.set_font("Helvetica", "B", 8)
    pdf.set_fill_color(241, 245, 249)
    pdf.cell(125, 6, "Variant / Artwork Description", border=1, fill=True)
    pdf.cell(35, 6, "Order Qty", border=1, align="R", fill=True)
    pdf.cell(30, 6, "Share", border=1, new_x="LMARGIN", new_y="NEXT", align="R", fill=True)

    pdf.set_font("Helvetica", "", 8)
    for v in data["Variant Details"]:
        pdf.set_x(pdf.l_margin)
        v_name = pdf_safe_text(v["variant"]) if v["variant"] else "Unnamed Label Design"
        pct = (v["quantity"] / data["Total Quantity"] * 100) if data["Total Quantity"] > 0 else 0
        pdf.cell(125, 6, v_name[:75], border=1)
        pdf.cell(35, 6, f"{v['quantity']:,}", border=1, align="R")
        pdf.cell(30, 6, f"{pct:.1f}%", border=1, new_x="LMARGIN", new_y="NEXT", align="R")
    pdf.ln(3)

    # 3. Technical & Print Specifications
    section_heading("3. Technical, Substrate & Proofing Specifications")
    key_value("Substrate Colour:", data["Substrate Colour"])
    key_value("Substrate Finish:", data["Substrate Type"])
    key_value("Material & Finishing:", data["Material Specs"])
    key_value("Colour Matching:", ", ".join(data["Colour Matching"]) if data["Colour Matching"] else "Standard CMYK")
    if "Pantone" in data["Colour Matching"]:
        key_value("Pantone Details:", f"{data['Pantone Qty']} label(s) - {data['Pantone Refs']}")
    if "Match to target supplied" in data["Colour Matching"]:
        key_value("Target Match Details:", f"{data['Target Qty']} label(s) - {data['Target Refs']}")
    key_value("Print Method:", f"Digital: {data['Digital Print']} | Litho: {data['Litho Print']}")
    key_value("Epson Proofs Required:", f"{data['Epson Proof Qty']} proof(s)")
    key_value("Layflat Supplied:", data["Layflat Supplied"])
    pdf.ln(3)

    # 4. Agreed Supplier Cost Summary (Unmarked Base Prices)
    section_heading("4. Order Value Summary (Excl. Client Markup)")
    pdf.set_x(pdf.l_margin)
    pdf.set_font("Helvetica", "B", 8)
    pdf.set_fill_color(241, 245, 249)
    pdf.cell(95, 6, "Item Description", border=1, fill=True)
    pdf.cell(35, 6, "Unit Price", border=1, align="R", fill=True)
    pdf.cell(25, 6, "Qty", border=1, align="R", fill=True)
    pdf.cell(35, 6, "Sub Total", border=1, new_x="LMARGIN", new_y="NEXT", align="R", fill=True)

    for item in data["Line Items"]:
        if item["qty"] > 0 or item["base_subtotal"] > 0:
            pdf.set_x(pdf.l_margin)
            pdf.set_font("Helvetica", "", 8)
            unit_fmt = f"R {item['base_unit_price']:,.4f}" if item["decimals"] == 4 else f"R {item['base_unit_price']:,.2f}"
            pdf.cell(95, 6, pdf_safe_text(item["description"])[:55], border=1)
            pdf.cell(35, 6, unit_fmt, border=1, align="R")
            pdf.cell(25, 6, f"{item['qty']:,}", border=1, align="R")
            pdf.cell(35, 6, f"R {item['base_subtotal']:,.2f}", border=1, new_x="LMARGIN", new_y="NEXT", align="R")

    pdf.ln(2)
    pdf.set_x(pdf.l_margin)
    pdf.set_font("Helvetica", "B", 9)
    pdf.cell(155, 6, "Net Order Subtotal:", border=0, align="R")
    pdf.cell(35, 6, f"R {data['Base Net Subtotal']:,.2f}", border=1, new_x="LMARGIN", new_y="NEXT", align="R")

    pdf.set_x(pdf.l_margin)
    pdf.cell(155, 6, "15% VAT:", border=0, align="R")
    pdf.cell(35, 6, f"R {data['Base VAT (15%)']:,.2f}", border=1, new_x="LMARGIN", new_y="NEXT", align="R")

    pdf.set_x(pdf.l_margin)
    pdf.set_fill_color(226, 232, 240)
    pdf.cell(155, 7, "Total Order Value (Incl. VAT):", border=0, align="R")
    pdf.cell(35, 7, f"R {data['Base Gross Total']:,.2f}", border=1, new_x="LMARGIN", new_y="NEXT", align="R", fill=True)

    return bytes(pdf.output())


# -----------------------------------------------------------------------------
# SESSION STATE HELPERS & CALLBACKS
# -----------------------------------------------------------------------------

def distribute_moq_evenly():
    """Distributes the selected dieline's MOQ evenly across the configured variants."""
    die_key = st.session_state.get("selected_dieline", DIELINE_OPTIONS[0])
    moq = DIELINE_SPECS[die_key]["moq"]
    n_vars = int(st.session_state.get("num_variants", 1))

    base_share = moq // n_vars
    remainder = moq % n_vars

    for idx in range(n_vars):
        st.session_state[f"v_qty_{idx}"] = base_share + (1 if idx < remainder else 0)


def on_substrate_colour_change():
    """Updates the Material Specifications description when toggling White vs Clear."""
    colour = st.session_state.get("substrate_colour", "White")
    if colour == "Clear":
        st.session_state["material_specs"] = CLEAR_MATERIAL_SPECS
    else:
        st.session_state["material_specs"] = WHITE_MATERIAL_SPECS


def apply_bulk_markup():
    """Applies a bulk markup percentage to whichever cost items are checked."""
    pct = float(st.session_state.get("bulk_markup_pct", 0.0))
    if st.session_state.get("bulk_apply_trial", False):
        st.session_state["trial_markup_pct"] = pct
    if st.session_state.get("bulk_apply_epson", False):
        st.session_state["epson_markup_pct"] = pct
    if st.session_state.get("bulk_apply_cm", False):
        st.session_state["cm_markup_pct"] = pct
    if st.session_state.get("bulk_apply_lf", False):
        st.session_state["lf_markup_pct"] = pct
    if st.session_state.get("bulk_apply_pt", False):
        st.session_state["pt_markup_pct"] = pct


def reset_all_form_fields():
    """Purges session state to reset all fields cleanly."""
    prefixes = ("v_name_", "v_qty_")
    known_keys = {
        "client_name", "selected_dieline", "num_variants", "substrate_colour",
        "substrate_type", "material_specs", "colour_match_type", "qty_pantone_match",
        "pantone_refs", "qty_target_match", "target_refs", "litho_print",
        "digital_print", "epson_proof_qty", "layflat_supplied", "cm_net", "cm_qty",
        "lf_net", "lf_qty", "pt_net", "pt_qty", "trial_markup_pct", "epson_markup_pct",
        "cm_markup_pct", "lf_markup_pct", "pt_markup_pct", "bulk_markup_pct",
        "po_number", "delivery_address", "order_notes"
    }
    for key in list(st.session_state.keys()):
        if key in known_keys or key.startswith(prefixes):
            del st.session_state[key]


# -----------------------------------------------------------------------------
# STREAMLIT UI
# -----------------------------------------------------------------------------

st.set_page_config(
    page_title="Bowler Packaging IML Digital Trial Quotation & Order Form", page_icon="🏷️", layout="wide"
)

# Initialize defaults before rendering
if "selected_dieline" not in st.session_state:
    st.session_state["selected_dieline"] = DIELINE_OPTIONS[0]
if "num_variants" not in st.session_state:
    st.session_state["num_variants"] = 1
if "substrate_colour" not in st.session_state:
    st.session_state["substrate_colour"] = "White"
if "material_specs" not in st.session_state:
    st.session_state["material_specs"] = WHITE_MATERIAL_SPECS

current_substrate = st.session_state["substrate_colour"]

st.title("🏷️ In-Mould Label (IML) Digital Trial Quotation & Order Builder")
st.markdown(
    "Configure your **Dieline & Variant Split**, apply optional **Client Markup %** on selected line items before exporting the **Client PDF Quote**, or **Convert directly to a Purchase Order for Ren-Flex**."
)

# Master Reference Table
with st.expander(f"📋 View Ren-Flex Digital Trial MOQ & Pricing Schedule ({current_substrate} Substrate)", expanded=False):
    ref_df = pd.DataFrame([
        {
            "Quote Ref": v["quote_ref"],
            "Die": v["die_code"],
            "Jar / Format": v["jar_name"],
            "Label Dimensions": v["size_mm"],
            "Trial MOQ (Total Run)": f"{v['moq']:,}",
            "White Unit Price": f"R {v['base_unit_price']:.4f}",
            "Clear Unit Price (+5%)": f"R {get_effective_unit_price(v['base_unit_price'], 'Clear'):.4f}",
            f"Supplier Run Cost ({current_substrate})": f"R {v['moq'] * get_effective_unit_price(v['base_unit_price'], current_substrate):,.2f}",
        }
        for v in DIELINE_SPECS.values()
    ])
    st.dataframe(ref_df, use_container_width=True, hide_index=True)

# Top Client Info & Refresh Bar
top_bar1, top_bar2 = st.columns([3.5, 1])
with top_bar1:
    client_name = st.text_input(
        "Client / Customer Name (for Client Quotation PDF)",
        key="client_name",
        placeholder="e.g., Bowler Plastics / End Brand Name",
    )
with top_bar2:
    st.write("")
    st.write("")
    st.button("🔄 Refresh / Clear Form", on_click=reset_all_form_fields, use_container_width=True)

st.divider()

# -----------------------------------------------------------------------------
# 1. SINGLE DIELINE SELECTION & VARIANT SPLIT
# -----------------------------------------------------------------------------
st.header("1. Dieline Selection & Variant Breakdown")

top_c1, top_c2, top_c3, top_c4 = st.columns([3, 1.3, 1.3, 1.4])
with top_c1:
    chosen_die = st.selectbox(
        "Select Dieline / Jar Size",
        options=DIELINE_OPTIONS,
        key="selected_dieline",
        on_change=distribute_moq_evenly,
    )
    spec = DIELINE_SPECS[chosen_die]
    st.caption(f"Ren-Flex Quote Ref: **{spec['quote_ref']}** | Label Size: **{spec['size_mm']}**")

with top_c2:
    n_vars = st.number_input(
        "No. of Variants in this Die",
        min_value=1,
        max_value=20,
        step=1,
        key="num_variants",
        on_change=distribute_moq_evenly,
    )

required_moq = spec["moq"]
base_unit_price = spec["base_unit_price"]
supplier_unit_price = get_effective_unit_price(base_unit_price, current_substrate)
supplier_run_charge = required_moq * supplier_unit_price

with top_c3:
    st.metric("Required MOQ (Total Qty)", f"{required_moq:,}")

with top_c4:
    price_note = (
        f"@ R {supplier_unit_price:.4f} / label (+5% Clear)"
        if current_substrate == "Clear"
        else f"@ R {supplier_unit_price:.4f} / label (White)"
    )
    st.metric(
        "Ren-Flex Base Run Cost",
        f"R {supplier_run_charge:,.2f}",
        delta=price_note,
        delta_color="normal" if current_substrate == "Clear" else "off"
    )

st.subheader("Allocate MOQ Across Label Designs / Variants")

if "v_qty_0" not in st.session_state:
    distribute_moq_evenly()

variant_details = []
allocated_sum = 0

for v in range(int(n_vars)):
    if f"v_qty_{v}" not in st.session_state:
        st.session_state[f"v_qty_{v}"] = 0

    vc1, vc2, vc3 = st.columns([3.5, 1.5, 1])
    with vc1:
        v_desc = st.text_input(
            f"Variant {v+1} Description / Artwork Name",
            key=f"v_name_{v}",
            placeholder=f'e.g., "{spec["jar_name"]} - Design {v+1}"',
        )
    with vc2:
        v_qty = st.number_input(
            f"Variant {v+1} Qty",
            min_value=0,
            max_value=required_moq,
            step=500,
            key=f"v_qty_{v}",
        )
    with vc3:
        pct_share = (v_qty / required_moq * 100) if required_moq > 0 else 0
        st.write("**% of MOQ**")
        st.write(f"{pct_share:.1f}%")

    allocated_sum += v_qty
    variant_details.append({
        "variant": v_desc,
        "quantity": v_qty,
    })

diff = required_moq - allocated_sum
is_moq_valid = (diff == 0)

val_c1, val_c2 = st.columns([3.5, 1.5])
with val_c1:
    if diff == 0:
        st.success(f"✅ **Allocated: {allocated_sum:,} / {required_moq:,} units** — Exact match to Dieline MOQ.")
    elif diff > 0:
        st.error(f"⚠️ **Allocated: {allocated_sum:,} / {required_moq:,} units** — Please allocate **{diff:,} more units** to reach the MOQ.")
    else:
        st.error(f"⚠️ **Allocated: {allocated_sum:,} / {required_moq:,} units** — Over MOQ by **{abs(diff):,} units**. Please reduce variant quantities.")
with val_c2:
    st.button(
        "⚖️ Split MOQ Evenly",
        on_click=distribute_moq_evenly,
        use_container_width=True,
    )

st.divider()

# -----------------------------------------------------------------------------
# 2. SUBSTRATE SPECIFICATIONS
# -----------------------------------------------------------------------------
st.header("2. Substrate Specifications")

col_sub1, col_sub2 = st.columns(2)
with col_sub1:
    substrate_colour = st.radio(
        "Substrate Colour (Clear adds +5% to Unit Price)",
        options=["White", "Clear"],
        horizontal=True,
        key="substrate_colour",
        on_change=on_substrate_colour_change,
    )
    if substrate_colour == "Clear":
        st.info("ℹ️ **Clear IML selected:** +5% markup applied to unit price and description updated to double-white printing.")

    substrate_type = st.radio(
        "Substrate Type",
        options=["Smooth", "Orange Peel"],
        horizontal=True,
        key="substrate_type"
    )

with col_sub2:
    material_specs = st.text_area(
        "Material Specifications",
        height=120,
        key="material_specs"
    )

st.divider()

# -----------------------------------------------------------------------------
# 3. COLOUR MATCHING & PRINT METHOD
# -----------------------------------------------------------------------------
st.header("3. Colour Matching & Print Method")

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
        f"Epson Proof (Qty @ R {EPSON_PROOF_UNIT_COST:,.2f} base cost each)",
        min_value=0,
        step=1,
        value=1,
        key="epson_proof_qty"
    )
    layflat_supplied = st.radio(
        "Layflat for layout and colour to be supplied?",
        options=["No", "Yes"],
        horizontal=True,
        key="layflat_supplied"
    )

st.divider()

# -----------------------------------------------------------------------------
# 4. FINANCIAL QUOTATION SUMMARY & SELECTIVE CLIENT MARKUP
# -----------------------------------------------------------------------------
st.header("4. Financial Summary & Client Markup")
st.caption(
    "Enter base supplier costs and optionally apply a **Markup %** on specific items for the **Client PDF Quote**. "
    "Converting to a **Ren-Flex Order** uses the unmarked supplier base costs."
)

# Optional Quick Bulk Markup Helper
with st.expander("⚡ Quick-Apply Markup % to Selected Items", expanded=False):
    bm1, bm2, bm3 = st.columns([1.5, 3, 1.5])
    with bm1:
        st.number_input("Markup % to Apply", min_value=0.0, max_value=500.0, value=15.0, step=1.0, key="bulk_markup_pct")
    with bm2:
        st.write("**Select items to mark up:**")
        chk1, chk2, chk3, chk4, chk5 = st.columns(5)
        chk1.checkbox("Trial Run", value=True, key="bulk_apply_trial")
        chk2.checkbox("Epson", value=False, key="bulk_apply_epson")
        chk3.checkbox("Colour", value=False, key="bulk_apply_cm")
        chk4.checkbox("Layflat", value=False, key="bulk_apply_lf")
        chk5.checkbox("Setup", value=False, key="bulk_apply_pt")
    with bm3:
        st.write("")
        st.button("Apply Markup %", on_click=apply_bulk_markup, use_container_width=True)

# Table Headers
h_desc, h_base, h_qty, h_basesub, h_mkp, h_clientsub = st.columns([2.8, 1.3, 1.0, 1.4, 1.2, 1.5])
h_desc.write("**Task Description**")
h_base.write("**Base Unit Cost**")
h_qty.write("**Qty**")
h_basesub.write("**Base Sub Total**")
h_mkp.write("**Markup %**")
h_clientsub.write("**Client Sub Total**")

# --- Row 1: Digital Trial Run ---
r1, r2, r3, r4, r5, r6 = st.columns([2.8, 1.3, 1.0, 1.4, 1.2, 1.5])
substrate_label = "IML Clear +5%" if substrate_colour == "Clear" else "IML White Gloss"
trial_desc = f"Digital Trial Run: {spec['die_code']} - {spec['jar_name']} ({substrate_label}, {int(n_vars)} variant(s))"
r1.write(f"**{trial_desc}**")
r2.write(f"R {supplier_unit_price:.4f}")
r3.write(f"{required_moq:,}")
r4.write(f"R {supplier_run_charge:,.2f}")
trial_markup = r5.number_input("Trial Markup %", min_value=0.0, max_value=500.0, value=0.0, step=1.0, key="trial_markup_pct", label_visibility="collapsed")
trial_client_unit = round(supplier_unit_price * (1 + trial_markup / 100.0), 4)
trial_client_sub = required_moq * trial_client_unit
r6.markdown(f"**R {trial_client_sub:,.2f}**")

# --- Row 2: Epson Proofs ---
e1, e2, e3, e4, e5, e6 = st.columns([2.8, 1.3, 1.0, 1.4, 1.2, 1.5])
e1.write("Epson Proofs")
e2.write(f"R {EPSON_PROOF_UNIT_COST:,.2f}")
e3.write(f"{int(epson_proof_qty)}")
epson_base_sub = int(epson_proof_qty) * EPSON_PROOF_UNIT_COST
e4.write(f"R {epson_base_sub:,.2f}")
epson_markup = e5.number_input("Epson Markup %", min_value=0.0, max_value=500.0, value=0.0, step=1.0, key="epson_markup_pct", label_visibility="collapsed")
epson_client_unit = round(EPSON_PROOF_UNIT_COST * (1 + epson_markup / 100.0), 2)
epson_client_sub = int(epson_proof_qty) * epson_client_unit
e6.markdown(f"**R {epson_client_sub:,.2f}**")

# --- Row 3: Colour Matching ---
c1, c2, c3, c4, c5, c6 = st.columns([2.8, 1.3, 1.0, 1.4, 1.2, 1.5])
c1.write("Colour Matching")
cm_net_val = c2.number_input("CM Base Unit", min_value=0.0, value=0.0, step=10.0, key="cm_net", label_visibility="collapsed")
cm_quantity = c3.number_input("CM Qty", min_value=0, value=0, step=1, key="cm_qty", label_visibility="collapsed")
cm_base_sub = cm_net_val * cm_quantity
c4.write(f"R {cm_base_sub:,.2f}")
cm_markup = c5.number_input("CM Markup %", min_value=0.0, max_value=500.0, value=0.0, step=1.0, key="cm_markup_pct", label_visibility="collapsed")
cm_client_unit = round(cm_net_val * (1 + cm_markup / 100.0), 2)
cm_client_sub = cm_quantity * cm_client_unit
c6.markdown(f"**R {cm_client_sub:,.2f}**")

# --- Row 4: Layflat ---
l1, l2, l3, l4, l5, l6 = st.columns([2.8, 1.3, 1.0, 1.4, 1.2, 1.5])
l1.write("Layflat")
lf_net_val = l2.number_input("LF Base Unit", min_value=0.0, value=0.0, step=10.0, key="lf_net", label_visibility="collapsed")
lf_quantity = l3.number_input("LF Qty", min_value=0, value=0, step=1, key="lf_qty", label_visibility="collapsed")
lf_base_sub = lf_net_val * lf_quantity
l4.write(f"R {lf_base_sub:,.2f}")
lf_markup = l5.number_input("LF Markup %", min_value=0.0, max_value=500.0, value=0.0, step=1.0, key="lf_markup_pct", label_visibility="collapsed")
lf_client_unit = round(lf_net_val * (1 + lf_markup / 100.0), 2)
lf_client_sub = lf_quantity * lf_client_unit
l6.markdown(f"**R {lf_client_sub:,.2f}**")

# --- Row 5: Additional Setup / Tooling ---
p1, p2, p3, p4, p5, p6 = st.columns([2.8, 1.3, 1.0, 1.4, 1.2, 1.5])
p1.write("Additional Setup / Tooling")
pt_net_val = p2.number_input("PT Base Unit", min_value=0.0, value=0.0, step=10.0, key="pt_net", label_visibility="collapsed")
pt_quantity = p3.number_input("PT Qty", min_value=0, value=0, step=1, key="pt_qty", label_visibility="collapsed")
pt_base_sub = pt_net_val * pt_quantity
p4.write(f"R {pt_base_sub:,.2f}")
pt_markup = p5.number_input("PT Markup %", min_value=0.0, max_value=500.0, value=0.0, step=1.0, key="pt_markup_pct", label_visibility="collapsed")
pt_client_unit = round(pt_net_val * (1 + pt_markup / 100.0), 2)
pt_client_sub = pt_quantity * pt_client_unit
p6.markdown(f"**R {pt_client_sub:,.2f}**")

st.divider()

# Totals Comparison (Base Supplier Cost vs. Client Marked-Up Quote)
base_net_subtotal = supplier_run_charge + epson_base_sub + cm_base_sub + lf_base_sub + pt_base_sub
base_vat_amount = base_net_subtotal * 0.15
base_gross_total = base_net_subtotal + base_vat_amount

client_net_subtotal = trial_client_sub + epson_client_sub + cm_client_sub + lf_client_sub + pt_client_sub
client_vat_amount = client_net_subtotal * 0.15
client_gross_total = client_net_subtotal + client_vat_amount
total_margin_zar = client_net_subtotal - base_net_subtotal

tot_col1, tot_col2, tot_col3 = st.columns([3, 3, 3])
with tot_col1:
    st.markdown("#### 🏭 Ren-Flex Supplier Cost")
    st.write(f"**Base Net Subtotal:** R {base_net_subtotal:,.2f}")
    st.write(f"**15% VAT:** R {base_vat_amount:,.2f}")
    st.write(f"**Supplier Gross Total:** R {base_gross_total:,.2f}")

with tot_col2:
    st.markdown("#### 📈 Applied Markup Margin")
    margin_pct_overall = (total_margin_zar / base_net_subtotal * 100) if base_net_subtotal > 0 else 0.0
    st.metric(
        "Total Net Markup",
        f"R {total_margin_zar:,.2f}",
        delta=f"{margin_pct_overall:.2f}% effective net margin",
        delta_color="normal" if total_margin_zar > 0 else "off"
    )

with tot_col3:
    st.markdown("#### 🧾 Client Quotation Total")
    st.write(f"**Client Net Subtotal:** R {client_net_subtotal:,.2f}")
    st.write(f"**15% VAT:** R {client_vat_amount:,.2f}")
    st.subheader(f"Gross Total: R {client_gross_total:,.2f}")

st.divider()

# -----------------------------------------------------------------------------
# 5. ACTIONS: SAVE CLIENT QUOTE PDF OR CONVERT TO REN-FLEX ORDER
# -----------------------------------------------------------------------------
st.header("5. Generate Client Quotation or Convert to Ren-Flex Order")

if not is_moq_valid:
    st.warning("⚠️ Please ensure the sum of variant quantities equals the Dieline MOQ before generating a quote or order.")

action_col1, action_col2 = st.columns(2)

# Build shared line items structure
line_items_payload = [
    {
        "description": trial_desc,
        "base_unit_price": supplier_unit_price,
        "client_unit_price": trial_client_unit,
        "qty": required_moq,
        "base_subtotal": supplier_run_charge,
        "markup_pct": trial_markup,
        "client_subtotal": trial_client_sub,
        "decimals": 4,
    },
    {
        "description": "Epson Proofs",
        "base_unit_price": EPSON_PROOF_UNIT_COST,
        "client_unit_price": epson_client_unit,
        "qty": int(epson_proof_qty),
        "base_subtotal": epson_base_sub,
        "markup_pct": epson_markup,
        "client_subtotal": epson_client_sub,
        "decimals": 2,
    },
    {
        "description": "Colour Matching",
        "base_unit_price": cm_net_val,
        "client_unit_price": cm_client_unit,
        "qty": int(cm_quantity),
        "base_subtotal": cm_base_sub,
        "markup_pct": cm_markup,
        "client_subtotal": cm_client_sub,
        "decimals": 2,
    },
    {
        "description": "Layflat",
        "base_unit_price": lf_net_val,
        "client_unit_price": lf_client_unit,
        "qty": int(lf_quantity),
        "base_subtotal": lf_base_sub,
        "markup_pct": lf_markup,
        "client_subtotal": lf_client_sub,
        "decimals": 2,
    },
    {
        "description": "Additional Setup / Tooling",
        "base_unit_price": pt_net_val,
        "client_unit_price": pt_client_unit,
        "qty": int(pt_quantity),
        "base_subtotal": pt_base_sub,
        "markup_pct": pt_markup,
        "client_subtotal": pt_client_sub,
        "decimals": 2,
    },
]

def build_payload(doc_type: str, po_number: str = "", delivery_addr: str = "", order_notes: str = "") -> dict:
    return {
        "Document Type": doc_type,
        "Client Name": client_name,
        "PO Number": po_number,
        "Delivery Address": delivery_addr,
        "Order Notes": order_notes,
        "Selected Dieline": chosen_die,
        "Quote Ref": spec["quote_ref"],
        "Die Code": spec["die_code"],
        "Jar Name": spec["jar_name"],
        "Dieline Size": spec["size_mm"],
        "Total Quantity": required_moq,
        "No of Variants": int(n_vars),
        "Variant Details": variant_details,
        "Substrate Colour": substrate_colour,
        "Substrate Type": substrate_type,
        "Material Specs": material_specs,
        "Colour Matching": colour_match_type,
        "Pantone Qty": qty_pantone_match,
        "Pantone Refs": pantone_refs,
        "Target Qty": qty_target_match,
        "Target Refs": target_refs,
        "Litho Print": litho_print,
        "Digital Print": digital_print,
        "Epson Proof Qty": int(epson_proof_qty),
        "Layflat Supplied": layflat_supplied,
        "Line Items": line_items_payload,
        "Base Net Subtotal": base_net_subtotal,
        "Base VAT (15%)": base_vat_amount,
        "Base Gross Total": base_gross_total,
        "Client Net Subtotal": client_net_subtotal,
        "Client VAT (15%)": client_vat_amount,
        "Client Gross Total": client_gross_total,
    }

# --- LEFT COLUMN: CLIENT QUOTATION WORKFLOW ---
with action_col1:
    with st.container(border=True):
        st.subheader("📄 Option A: Save & Export Client Quotation")
        st.write("Generates the client-facing PDF quotation using your **Marked-Up Client Prices** and logs it to Google Sheets.")
        st.write("")
        save_client_quote_btn = st.button(
            "Save Quotation & Generate Client PDF",
            type="primary",
            use_container_width=True,
            disabled=not is_moq_valid,
        )

        if save_client_quote_btn and is_moq_valid:
            payload = build_payload("Client Quotation")
            with st.spinner("Saving quotation and generating Client PDF..."):
                sheet_ok = append_to_google_sheets(payload)
                if sheet_ok:
                    st.info("✅ Client quotation recorded in Google Sheets.")
                pdf_bytes = generate_client_pdf_quote(payload)
                if pdf_bytes and HAS_FPDF:
                    st.success("🎉 Client Quotation PDF is ready!")
                    st.download_button(
                        label="📥 Download Client Quotation PDF",
                        data=pdf_bytes,
                        file_name=f"Client_IML_Quote_{spec['die_code'].replace(' ', '')}_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                    )

# --- RIGHT COLUMN: CONVERT TO REN-FLEX ORDER WORKFLOW ---
with action_col2:
    with st.container(border=True):
        st.subheader("🏭 Option B: Convert to Order for Ren-Flex")
        st.write("Converts this quote into a **Purchase Order for Ren-Flex Packaging** using base supplier costs.")
        po_number = st.text_input(
            "Purchase Order (PO) Number",
            value=f"PO-IML-{datetime.now().strftime('%Y%m%d-%H%M')}",
            key="po_number",
        )
        delivery_address = st.text_input(
            "Delivery Address / Dispatch Instruction",
            value="Delivered to 1 Address (Standard)",
            key="delivery_address",
        )
        order_notes = st.text_input(
            "Artwork / Production Instructions for Ren-Flex",
            placeholder="e.g., Print-ready PDFs attached via email; confirm Epson proof dispatch",
            key="order_notes",
        )
        convert_order_btn = st.button(
            "Convert to Order for Ren-Flex & Generate PO PDF",
            use_container_width=True,
            disabled=not is_moq_valid,
        )

        if convert_order_btn and is_moq_valid:
            if not po_number.strip():
                st.error("Please enter a Purchase Order (PO) Number before converting to an order.")
            else:
                order_payload = build_payload(
                    "Ren-Flex Order",
                    po_number=po_number.strip(),
                    delivery_addr=delivery_address.strip(),
                    order_notes=order_notes.strip(),
                )
                with st.spinner("Converting to Ren-Flex Purchase Order..."):
                    sheet_ok = append_to_google_sheets(order_payload)
                    if sheet_ok:
                        st.info(f"✅ Order ({po_number}) logged in Google Sheets.")
                    po_pdf_bytes = generate_renflex_order_pdf(order_payload)
                    if po_pdf_bytes and HAS_FPDF:
                        st.success(f"🎉 Ren-Flex Purchase Order ({po_number}) is ready!")
                        st.download_button(
                            label="📥 Download Ren-Flex Purchase Order PDF",
                            data=po_pdf_bytes,
                            file_name=f"RenFlex_Order_{po_number.replace(' ', '_')}.pdf",
                            mime="application/pdf",
                            use_container_width=True,
                        )
