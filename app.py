import streamlit as st

from pinyin_pptx import add_pinyin

st.set_page_config(page_title="歌詞拼音工具", page_icon="🎵")
st.title("🎵 歌詞拼音工具")
st.caption("上傳歌詞 PPTX,自動在中文歌詞上方加上漢語拼音,格式與背景維持不變。")

files = st.file_uploader("上傳 PPTX 檔案(可多選)", type="pptx", accept_multiple_files=True)

with st.expander("進階設定"):
    min_pt = st.number_input(
        "只處理字級大於等於(pt)的中文行", value=40, min_value=8, max_value=96,
        help="用來略過標題、footer、底部小字幕。預設 40 只會處理主歌詞。",
    )
    ratio = st.slider("拼音字級比例(相對中文)", 0.3, 0.8, 0.45, 0.05)

if files:
    for f in files:
        try:
            with st.spinner(f"處理中:{f.name}"):
                out = add_pinyin(f, min_pt=min_pt, ratio=ratio)
            new_name = f.name.rsplit(".", 1)[0] + "_拼音.pptx"
            st.success(f"完成:{f.name}")
            st.download_button(
                f"⬇️ 下載 {new_name}",
                out,
                file_name=new_name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                key=f.name,
            )
        except Exception as e:
            st.error(f"{f.name} 處理失敗:{e}")

st.divider()
st.caption("讀音修正:「祢 → nǐ」「尊主為大 → wéi」已內建;發現其他錯誤讀音請回報管理員。")
