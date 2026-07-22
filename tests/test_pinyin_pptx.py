"""Tests for pinyin_pptx: existing-pinyin removal + insertion below lyrics."""
import io

import pytest
from pptx import Presentation
from pptx.util import Inches, Pt

from pinyin_pptx import add_pinyin, is_pinyin_line


def make_pptx(lines, size_pt=40):
    """One slide, one textbox; each item in `lines` becomes a paragraph."""
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(4))
    tf = box.text_frame
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        run = p.add_run()
        run.text = line
        run.font.size = Pt(size_pt)
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf


def all_para_texts(buf):
    prs = Presentation(buf)
    texts = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    texts.append("".join(r.text for r in p.runs))
    return texts


class TestIsPinyinLine:
    def test_detects_tone_marked_pinyin(self):
        assert is_pinyin_line("zūn zhǔ wéi dà")
        assert is_pinyin_line("  yē sū ài nǐ  ")

    def test_rejects_chinese(self):
        assert not is_pinyin_line("尊主為大")
        assert not is_pinyin_line("zūn 主")

    def test_rejects_plain_english(self):
        assert not is_pinyin_line("Jesus loves me")
        assert not is_pinyin_line("Amazing grace")

    def test_rejects_empty(self):
        assert not is_pinyin_line("")
        assert not is_pinyin_line("   ")


class TestAddPinyin:
    def test_pinyin_added_below_lyric(self):
        out = add_pinyin(make_pptx(["尊主為大"]))
        texts = all_para_texts(out)
        assert texts[0] == "尊主為大"
        assert "zūn" in texts[1] and "dà" in texts[1]

    def test_existing_pinyin_above_is_removed(self):
        out = add_pinyin(make_pptx(["zūn zhǔ wéi dà", "尊主為大"]))
        texts = [t for t in all_para_texts(out) if t.strip()]
        assert texts[0] == "尊主為大"          # 拼音不再出現在歌詞上面
        assert "zūn" in texts[1]               # 新拼音在歌詞下面
        assert len(texts) == 2

    def test_idempotent_rerun(self):
        once = add_pinyin(make_pptx(["尊主為大"]))
        twice = add_pinyin(once)
        texts = [t for t in all_para_texts(twice) if t.strip()]
        assert len(texts) == 2

    def test_english_lines_untouched(self):
        out = add_pinyin(make_pptx(["Jesus loves me", "尊主為大"]))
        texts = [t for t in all_para_texts(out) if t.strip()]
        assert texts[0] == "Jesus loves me"
        assert texts[1] == "尊主為大"

    def test_small_font_skipped(self):
        out = add_pinyin(make_pptx(["尊主為大"], size_pt=20), min_pt=40)
        texts = [t for t in all_para_texts(out) if t.strip()]
        assert texts == ["尊主為大"]

    def test_pinyin_only_textbox_is_blanked(self):
        out = add_pinyin(make_pptx(["zūn zhǔ wéi dà"]))
        texts = [t for t in all_para_texts(out) if t.strip()]
        assert texts == []
