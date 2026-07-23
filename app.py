import io

import streamlit as st
import streamlit.components.v1 as components

from pinyin_pptx import PINYIN_VERSION, add_pinyin

MIME_PPTX = "application/vnd.openxmlformats-officedocument.presentationml.presentation"


@st.cache_data(show_spinner=False, max_entries=20)
def process_pptx(data: bytes, min_pt: float, pinyin_pt: float,
                 version: int = PINYIN_VERSION) -> bytes:
    """Cached: identical file + settings reuse the previous result, so UI
    reruns don't re-scan the deck. `version` keys the cache to the core
    logic, so releases invalidate stale results."""
    return add_pinyin(io.BytesIO(data), min_pt=min_pt, pinyin_pt=pinyin_pt).getvalue()

st.set_page_config(page_title="詩歌拼音 Shīgē Pinyin", page_icon="🎵", layout="centered")

st.markdown(r"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sora:wght@500;600;700&family=Inter:wght@400;500;600&display=swap');

:root{
  --paper:#FBFAF7; --surface:#FFFFFF; --ink:#17161C; --muted:#6B7280;
  --line:#E7E5DF; --accent:#0C8A78; --accent-strong:#0A7365; --accent-tint:#E6F4F1;
}

/* hide Streamlit chrome for a cleaner product surface */
#MainMenu, header[data-testid="stHeader"], footer, [data-testid="stToolbar"]{display:none!important;}
/* hide Community Cloud creator avatar + Streamlit badge (bottom-right) */
[class*="viewerBadge"], [class*="profileContainer"], [class*="profilePreview"],
[data-testid="appCreatorAvatar"], a[href*="streamlit.io/cloud"],
a[href*="share.streamlit.io/user"]{ display:none!important; }

.stApp{ background:var(--paper); }
.block-container{
  max-width:660px; margin-left:auto!important; margin-right:auto!important;
  padding-top:3.2rem; padding-bottom:4rem;
}
html, body, [class*="css"]{ font-family:'Inter',system-ui,sans-serif; color:var(--ink); }

/* ---------- hero ---------- */
.hero{ text-align:center; margin-bottom:2.4rem; }
.hero *, .hero p, .hero h1{ text-align:center!important; }
.eyebrow{
  font-size:.72rem; letter-spacing:.18em; text-transform:uppercase;
  color:var(--accent-strong); font-weight:600; margin-bottom:1rem;
}
.wordmark{
  font-family:'Sora',sans-serif; font-weight:700; font-size:2.9rem;
  line-height:1.15; letter-spacing:-.01em; margin:0; color:var(--ink)!important;
  display:flex; align-items:flex-start; justify-content:center!important; gap:.12em;
  width:100%;
}
/* character + pinyin stacked below — mirrors the product output */
.wordmark .zi{
  display:inline-flex; flex-direction:column; align-items:center; font-weight:700;
}
/* brand gradient on the big glyphs — echoes the download button */
.wordmark .zi, .wordmark .latin{
  background:linear-gradient(160deg,#12A38B 0%,#0A7365 70%);
  -webkit-background-clip:text; background-clip:text;
  -webkit-text-fill-color:transparent; color:var(--accent-strong);
}
.wordmark .py{
  font-family:'Inter',sans-serif; font-size:.32em; font-weight:600;
  color:var(--accent); -webkit-text-fill-color:var(--accent);
  letter-spacing:.02em; margin-top:.18em;
}
.wordmark .latin{ margin-left:.18em; line-height:1.15; }
/* Streamlit appends a hidden anchor-link element inside headings —
   it takes up flex space after "Pinyin" and skews centering */
.wordmark a, .wordmark svg,
.wordmark [data-testid="stHeaderActionElements"]{ display:none!important; }
.subtitle{
  margin:1.1rem auto 0!important; max-width:34rem; color:var(--muted);
  font-size:1.02rem; line-height:1.6; text-align:center!important;
}

/* ---------- file uploader → drop zone ---------- */
[data-testid="stFileUploader"]{ margin-top:.4rem; }
[data-testid="stFileUploader"] label{ display:none; }
[data-testid="stFileUploaderDropzone"]{
  display:flex!important; flex-direction:column; align-items:center; justify-content:center;
  gap:.55rem; background:var(--surface); border:2px dashed #CFCEC7; border-radius:18px;
  padding:2.7rem 1.5rem; transition:all .18s ease; cursor:pointer;
}
[data-testid="stFileUploaderDropzone"]:hover{
  border-color:var(--accent); background:var(--accent-tint);
}
[data-testid="stFileUploaderDropzoneInstructions"]{ display:none!important; }
/* explicit flex order → icon, then text, then the Browse button */
[data-testid="stFileUploaderDropzone"]::before{
  order:0; content:"↑"; font-size:1.7rem; line-height:1; color:var(--accent);
  font-weight:700;
  width:3rem; height:3rem; display:flex; align-items:center; justify-content:center;
  background:var(--accent-tint); border-radius:50%;
}
[data-testid="stFileUploaderDropzone"]:hover::before{ background:#fff; }
[data-testid="stFileUploaderDropzone"]::after{
  order:1; content:"Drag & drop your PPTX here\A or click to browse";
  white-space:pre-line; text-align:center; color:var(--ink);
  font-weight:600; font-size:1.02rem; line-height:1.5;
}
/* the built-in Browse button, restyled as a quiet pill */
[data-testid="stFileUploaderDropzone"] button{
  order:2; margin-top:.35rem; background:transparent!important; color:var(--accent-strong)!important;
  border:1.5px solid var(--accent)!important; border-radius:10px!important;
  font-weight:600!important; padding:.4rem 1.1rem!important; transition:all .15s ease;
}
[data-testid="stFileUploaderDropzone"] button:hover{
  background:var(--accent)!important; color:#fff!important;
}
/* uploaded-file chips */
[data-testid="stFileUploaderFile"]{
  background:var(--surface); border:1px solid var(--line); border-radius:10px;
  padding:.5rem .7rem; margin-top:.5rem;
}

/* ---------- advanced settings expander ---------- */
[data-testid="stExpander"]{ border:none!important; margin-top:1rem; }
[data-testid="stExpander"] details{
  background:var(--accent-tint); border:1px solid #CFE7E1!important; border-radius:12px;
}
[data-testid="stExpander"] summary{ font-weight:500; color:var(--accent-strong); }
[data-testid="stExpander"] summary:hover{ color:var(--accent-strong); }

/* number inputs */
[data-testid="stNumberInput"] input{ border-radius:9px; }
[data-testid="stNumberInput"] label p{ font-weight:500; color:var(--ink); }

/* ---------- result card ---------- */
[data-testid="stVerticalBlockBorderWrapper"]:has(.result-head){
  background:var(--surface); border:1px solid var(--line)!important;
  border-radius:16px; padding:.35rem .35rem;
  box-shadow:0 1px 2px rgba(20,18,30,.04);
}
.result-head{
  display:flex; align-items:center; gap:.6rem; padding:.35rem .35rem .1rem;
}
.result-head .check{
  flex:none; width:1.5rem; height:1.5rem; border-radius:50%;
  background:var(--accent); color:#fff; font-size:.85rem; font-weight:700;
  display:flex; align-items:center; justify-content:center;
}
.result-head .fname{
  font-weight:600; color:var(--ink); font-size:.98rem; overflow:hidden;
  text-overflow:ellipsis; white-space:nowrap;
}
.result-head .pill{
  margin-left:auto; flex:none; font-size:.7rem; font-weight:600; letter-spacing:.04em;
  text-transform:uppercase; color:var(--accent-strong);
  background:var(--accent-tint); padding:.2rem .55rem; border-radius:999px;
}

/* ---------- download button (the prominent action) ---------- */
[data-testid="stDownloadButton"] button{
  width:100%; border:none!important; border-radius:12px!important;
  background:linear-gradient(135deg,#0FA18C,#0A7365)!important; color:#fff!important;
  font-family:'Inter',sans-serif; font-weight:600!important; font-size:1.02rem!important;
  padding:.85rem 1.2rem!important; letter-spacing:.01em;
  box-shadow:0 6px 16px rgba(12,138,120,.28)!important;
  transition:transform .15s ease, box-shadow .15s ease, filter .15s ease;
}
[data-testid="stDownloadButton"] button:hover{
  transform:translateY(-1px); filter:brightness(1.04);
  box-shadow:0 10px 22px rgba(12,138,120,.34)!important;
}
[data-testid="stDownloadButton"] button:active{ transform:translateY(0); }
[data-testid="stDownloadButton"] button p::before{ content:"↓  "; font-weight:700; }

/* alerts */
[data-testid="stAlert"]{ border-radius:12px; }

/* footer note */
.foot{
  margin-top:2.6rem; padding-top:1.3rem; border-top:1px solid var(--line);
  color:var(--muted); font-size:.82rem; line-height:1.6; text-align:center;
}
.foot code{ background:var(--accent-tint); color:var(--accent-strong);
  padding:.05rem .35rem; border-radius:5px; font-size:.9em; }

/* accessibility: focus + reduced motion */
:focus-visible{ outline:2px solid var(--accent); outline-offset:2px; }
@media (prefers-reduced-motion:reduce){ *{ transition:none!important; } }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
  <div class="eyebrow">Pinyin for worship slides</div>
  <h1 class="wordmark">
    <span class="zi">詩<span class="py">shī</span></span>
    <span class="zi">歌<span class="py">gē</span></span>
    <span class="latin">Pinyin</span>
  </h1>
  <p class="subtitle">Add pinyin below your Chinese lyrics.</p>
</div>
""", unsafe_allow_html=True)

files = st.file_uploader("Upload PPTX", type="pptx", accept_multiple_files=True)

with st.expander("Advanced settings", expanded=True):
    min_pt = st.number_input(
        "Minimum Chinese font size to process (pt)",
        value=40, min_value=8, max_value=96,
        help="Lines smaller than this are skipped — keeps titles, footers, and "
             "small subtitles clean. Lower it if your lyrics are smaller than 40pt.",
    )
    st.caption(f"Only Chinese text {min_pt} pt or larger gets pinyin — "
               "smaller text like titles and footers is left untouched.")
    pinyin_pt = st.number_input(
        "Pinyin font size (pt)", value=20, min_value=6, max_value=60,
        help="Absolute size of the pinyin text.",
    )

if files:
    for f in files:
        try:
            with st.spinner(f"Adding pinyin to {f.name}…"):
                data = process_pptx(f.getvalue(), min_pt, pinyin_pt)
            new_name = f.name.rsplit(".", 1)[0] + "_pinyin.pptx"
            with st.container(border=True):
                st.markdown(
                    f'<div class="result-head">'
                    f'<span class="check">✓</span>'
                    f'<span class="fname">{new_name}</span>'
                    f'<span class="pill">Ready</span></div>',
                    unsafe_allow_html=True,
                )
                st.download_button(
                    f"Download {new_name}",
                    data,
                    file_name=new_name,
                    mime=MIME_PPTX,
                    key=f.name,
                )
            # done: notify and scroll to the download button
            st.toast(f"✅ Pinyin added — {new_name} is ready", icon="🎵")
            components.html("""
            <script>
              const btns = window.parent.document
                  .querySelectorAll('[data-testid="stDownloadButton"]');
              if (btns.length) btns[btns.length - 1]
                  .scrollIntoView({behavior: 'smooth', block: 'center'});
            </script>""", height=0)
        except Exception as e:
            st.error(f"Couldn't process {f.name}: {e}")

st.markdown("""
<div class="foot">
  Spotted a wrong reading? Let the administrator know.
</div>
""", unsafe_allow_html=True)
