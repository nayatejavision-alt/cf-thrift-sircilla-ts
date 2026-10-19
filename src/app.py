import streamlit as st
import pandas as pd
from parser import parse_term_deposit_file

st.set_page_config(page_title="టర్మ్ & రికరింగ్ డిపాజిట్స్ డాష్‌బోర్డ్", layout="wide")

st.title("🏦 టర్మ్ & రికరింగ్ డిపాజిట్స్ డాష్‌బోర్డ్")
st.write("బ్యాంకు నుండి పొందిన డిపాజిట్ వివరాల ఫైల్‌ను కింద అప్‌లోడ్ చేయండి.")

uploaded_file = st.file_uploader("టెక్స్ట్ ఫైల్‌ను ఎంచుకోండి", type=["txt"])

if uploaded_file is not None:
    with open("temp_file.txt", "wb") as f:
        f.write(uploaded_file.getbuffer())

    df = parse_term_deposit_file("temp_file.txt")

    if not df.empty:
        st.success(f" విజయవంతంగా డేటా లోడ్ అయింది! మొత్తం రికార్డులు: {len(df)}")

        # ముఖ్యమైన గణాంకాలు
        total_balance = df['Balance'].sum()
        st.metric(label="మొత్తం డిపాజిట్ బ్యాలెన్స్", value=f"₹ {total_balance:,.2f}")

        # సెక్షన్ల వారీగా సారాంశం
        st.subheader("📊 సెక్షన్ల వారీగా వివరాలు")
        summary = df.groupby('Section')['Balance'].agg(['count', 'sum']).reset_index()
        summary.columns = ['సెక్షన్', 'ఖాతాల సంఖ్య', 'మొత్తం బ్యాలెన్స్ (₹)']
        st.dataframe(summary, use_container_width=True)

        # డేటా టేబుల్ చూపించడం
        st.subheader("📋 అన్ని ఖాతాల వివరాలు")
        st.dataframe(df, use_container_width=True)
    else:
        st.error("ఫైల్ నుండి డేటాను చదవడం సాధ్యపడలేదు. దయచేసి సరైన ఫైల్ ఇవ్వండి.")
