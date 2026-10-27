import streamlit as st

st.set_page_config(
    page_title="IML Label Quotation Form", page_icon="🏷️", layout="wide"
)

st.title("🏷️ In-Mould Label (IML) Quotation Request")
st.markdown(
    "Please fill out the details below to generate an accurate quotation for your IML label range."
)

with st.form("iml_quote_form"):

    # -------------------------------------------------------------------------
    # 1. QUANTITIES & VARIANTS
    # -------------------------------------------------------------------------
    st.header("1. Quantities & Variant Breakdown")

    col1, col2 = st.columns(2)
    with col1:
        total_qty = st.number_input(
            "Total Qty of IML Labels in Range", min_value=1, step=100, value=10000
        )
    with col2:
        num_variants = st.number_input(
            "No. of IML Label Variants in Range", min_value=1, step=1, value=1
        )

    st.subheader("Variant Quantities & Descriptions")
    st.info(
        "Specify the details and quantities for each individual variant in the range."
    )

    variant_details = []
    # Dynamic textboxes generated based on the number of variants selected
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

    # -------------------------------------------------------------------------
    # 2. SUBSTRATE SPECIFICATIONS
    # -------------------------------------------------------------------------
    st.header("2. Substrate Specifications")

    col_sub1, col_sub2 = st.columns(2)
    with col_sub1:
        substrate_colour = st.radio(
            "Substrate Colour",
            options=["White", "Clear"],
            horizontal=True,
        )
        substrate_type = st.radio(
            "Substrate Type",
            options=["Smooth", "Orange Peel"],
            horizontal=True,
        )

    with col_sub2:
        material_specs = st.text_area(
            "Material Specifications",
            placeholder="Specify film thickness, barrier properties, or specific resin grade requirements...",
            height=120,
        )

    st.divider()

    # -------------------------------------------------------------------------
    # 3. JAR COMPATIBILITY
    # -------------------------------------------------------------------------
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
        "Select Target Jars (Specify)", options=jar_options
    )

    st.divider()

    # -------------------------------------------------------------------------
    # 4. COLOUR MATCHING & PROCESS
    # -------------------------------------------------------------------------
    st.header("4. Colour Matching & Print Method")

    col_c1, col_c2 = st.columns(2)
    with col_c1:
        colour_match_type = st.multiselect(
            "Colour Matching Method Required",
            options=["Pantone", "Match to target supplied"],
        )

        qty_pantone_match = 0
        pantone_refs = ""
        if "Pantone" in colour_match_type:
            qty_pantone_match = st.number_input(
                "No. of labels to match to Pantone (Qty)", min_value=0, step=1
            )
            pantone_refs = st.text_input(
                "Specify Pantone References",
                placeholder='e.g., "PMS 185 C, PMS 286 C"',
            )

        qty_target_match = 0
        target_refs = ""
        if "Match to target supplied" in colour_match_type:
            qty_target_match = st.number_input(
                "No. of labels to match to Target (Qty)", min_value=0, step=1
            )
            target_refs = st.text_input(
                "Specify Target References / Samples",
                placeholder='e.g., "Physical sample ref #1042"',
            )

    with col_c2:
        st.write("**Print Process & Proofing**")
        litho_print = st.radio(
            "Litho Print Required?", options=["No", "Yes"], horizontal=True
        )
        digital_print = st.radio(
            "Digital Print Required?", options=["No", "Yes"], horizontal=True
        )

        epson_proof_qty = st.number_input(
            "Epson Proof (Qty)", min_value=0, step=1, value=1
        )
        layflat_supplied = st.radio(
            "Layflat for layout and colour to be supplied?",
            options=["No", "Yes"],
            horizontal=True,
        )

    st.divider()

    # -------------------------------------------------------------------------
    # 5. TRIALS
    # -------------------------------------------------------------------------
    st.header("5. Production Trials")

    trials_required = st.radio(
        "Trials Required?", options=["No", "Yes"], horizontal=True
    )

    trial_variants = 0
    trial_qty_per_variant = 0

    if trials_required == "Yes":
        col_t1, col_t2 = st.columns(2)
        with col_t1:
            trial_variants = st.number_input(
                "Trials (No. of variants)", min_value=1, step=1, value=1
            )
        with col_t2:
            trial_qty_per_variant = st.number_input(
                "Trials (Quantity per variant)",
                min_value=1,
                step=100,
                value=500,
            )

    st.divider()

    # -------------------------------------------------------------------------
    # SUBMIT BUTTON
    # -------------------------------------------------------------------------
    submit_button = st.form_submit_button(
        "Submit Quotation Request", use_container_width=True
    )

if submit_button:
    st.success("Quotation request submitted successfully!")

    # Summary Output Dictionary
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
    }

    st.subheader("Submitted Quotation Summary")
    st.json(quote_data)