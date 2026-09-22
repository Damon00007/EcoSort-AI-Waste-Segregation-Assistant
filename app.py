import streamlit as st

st.set_page_config(page_title="EcoSort AI", page_icon="♻️", layout="centered")

st.title("♻️ EcoSort AI – Smart Waste Segregation Assistant")
st.caption("1M1B AI for Sustainability Virtual Internship | Aligned with UN SDG 12")

waste_kb = {
    "plastic bottle": {
        "bin": "🔵 Blue Bin (Dry Recyclable)",
        "prep": "Empty leftover liquid, lightly rinse, crush bottle, and put cap back on.",
        "impact": "Saves landfill space and reduces raw petroleum consumption for virgin plastic."
    },
    "banana peel": {
        "bin": "🟢 Green Bin (Wet / Organic Waste)",
        "prep": "Put directly into compost/organic bin. Do not enclose in a plastic bag.",
        "impact": "Naturally converts to nutrient-rich compost and prevents methane buildup."
    },
    "paper cup": {
        "bin": "🔵 Blue Bin (Dry / Non-recyclable Co-processing)",
        "prep": "Empty remaining drink. Plastic waterproof coating prevents normal paper recycling.",
        "impact": "Directs laminated paper to controlled waste-to-energy or RDF plants."
    },
    "lithium battery": {
        "bin": "🔴 Red / Black Bin (Hazardous E-Waste)",
        "prep": "Cover terminals with tape and hand over to campus or municipal e-waste center.",
        "impact": "Prevents toxic chemicals (lithium, cobalt) from leaching into groundwater."
    }
}

item = st.text_input("Enter waste item:", placeholder="e.g. plastic bottle, banana peel, lithium battery").lower().strip()

if st.button("Segregate & Advise"):
    if not item:
        st.error("Please enter an item name.")
    else:
        found = False
        for key, val in waste_kb.items():
            if key in item or item in key:
                st.success(f"**Recommended Bin:** {val['bin']}")
                st.info(f"**Action Required:** {val['prep']}")
                st.write(f"🌱 **Sustainability Benefit:** {val['impact']}")
                found = True
                break
        if not found:
            st.warning("General Guideline: Food/Wet goes to Green Bin, Clean Dry Plastic/Paper to Blue Bin, Electronics/Hazardous to Red/Black Bin.")