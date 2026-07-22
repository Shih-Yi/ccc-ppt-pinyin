import streamlit as st

from pinyin_pptx import add_pinyin

st.set_page_config(page_title="Lyrics Pinyin Tool", page_icon="🎵")
st.title("🎵 Lyrics Pinyin Tool")
st.caption("Upload a lyrics PPTX to automatically add Hanyu Pinyin below each Chinese lyric line. Formatting and background stay unchanged.")

files = st.file_uploader("Upload PPTX file(s)", type="pptx", accept_multiple_files=True)

with st.expander("Advanced settings"):
    min_pt = st.number_input(
        "Only process Chinese lines at or above this font size (pt)",
        value=40, min_value=8, max_value=96,
        help="Skips titles, footers, and small bottom subtitles. Default 40 processes only the main lyrics.",
    )
    pinyin_pt = st.number_input(
        "Pinyin font size (pt)", value=20, min_value=6, max_value=60,
        help="Absolute font size for the pinyin line. Default 20pt.",
    )

if files:
    for f in files:
        try:
            with st.spinner(f"Processing: {f.name}"):
                out = add_pinyin(f, min_pt=min_pt, pinyin_pt=pinyin_pt)
            new_name = f.name.rsplit(".", 1)[0] + "_pinyin.pptx"
            st.success(f"Done: {f.name}")
            st.download_button(
                f"⬇️ Download {new_name}",
                out,
                file_name=new_name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                key=f.name,
            )
        except Exception as e:
            st.error(f"{f.name} failed: {e}")

st.divider()
st.caption("Built-in pronunciation fixes: 祢 → nǐ, 尊主為大 → wéi. Report any other incorrect readings to the administrator.")
