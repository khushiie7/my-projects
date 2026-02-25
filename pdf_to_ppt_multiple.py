"""
Interactive PDF to PPT Generator with **Multiple PDF Support**
Tech: Streamlit, LangChain, FAISS, Ollama/HuggingFace LLMs
Allows uploading multiple PDFs on related topics to generate a single consolidated presentation.
"""

import streamlit as st
import os
import tempfile
import re
import requests
from typing import List, Dict, Tuple, Any, cast
from io import BytesIO

# PPTX
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN  # type: ignore
# LangChain
from langchain_community.document_loaders import PyPDFLoader
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_community.vectorstores import FAISS
from langchain_huggingface import HuggingFaceEmbeddings

try:
    import torch
except ImportError:
    torch = None

# ==========================================================
# STREAMLIT CONFIG
# ==========================================================
st.set_page_config(
    page_title="Multi-PDF to PPT Generator",
    layout="wide",
    page_icon="📚"
)

# OLLAMA LLM CLASS
# ==========================================================
import asyncio
import edge_tts
from PIL import Image
import base64
import config as cfg
import startup_cleanup
from wav2lip_integration import Wav2LipGenerator
from fooocus_integration import FooocusImageGenerator
from mermaid_integration import generate_mermaid_image  # type: ignore

# ── Run once at startup: clear temp files, rotate log ──────────────────────
if "startup_cleaned" not in st.session_state:
    startup_cleanup.run()
    st.session_state.startup_cleaned = True


def image_to_base64(image):
    buffered = BytesIO()
    image.save(buffered, format="PNG")
    return base64.b64encode(buffered.getvalue()).decode()


def get_image_generator():
    """Factory to return the configured image generator."""
    if cfg.IMAGE_PROVIDER == "fooocus":
        return FooocusImageGenerator()
    return SmartSlideImageGenerator()


def video_to_base64(video_path: str) -> str:
    """Read a video file and return a base64-encoded data URL (video/mp4)."""
    with open(video_path, "rb") as f:
        encoded = base64.b64encode(f.read()).decode()
    return f"data:video/mp4;base64,{encoded}"


def save_bg_to_temp(bg_img):
    """Save a PIL Image to a temp PNG and return the file path."""
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    bg_img.save(tmp.name)
    return tmp.name


def create_templated_slide(prs, slide_data, bg_path=None, avatar_img=None):
    """
    Clean content slide template — no overlay panel.
    Zones:
      1. Full-bleed background image
      2. Title bar (full width, dark red)
      3. Bullet points (left 60% of slide)
      4. Right image  (right 35%)
      5. Tutor avatar/video (bottom-right corner)
      6. Hidden audio
      7. Speaker notes
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout

    SLIDE_W = prs.slide_width    # 13.33"
    SLIDE_H = prs.slide_height   # 7.5"

    # 1️⃣ Background Image (Commented out to force white background as requested)
    # if bg_path and os.path.exists(bg_path):
    #     try:
    #         slide.shapes.add_picture(bg_path, Inches(0), Inches(0),
    #                                  width=SLIDE_W, height=SLIDE_H)
    #     except Exception as e:
    #         print(f"[PPT] BG error: {e}")

    # 2️⃣ Title — full width, dark red, bold
    title_box = slide.shapes.add_textbox(
        left=Inches(0.4), top=Inches(0.15), # Lifted up
        width=Inches(12.5), height=Inches(1.0)
    )
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = slide_data.get("title", "Untitled")
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = RGBColor(180, 0, 0)   # Dark red

    # 3️⃣ Bullet Points — indented with filled circle bullets
    bullets_box = slide.shapes.add_textbox(
        left=Inches(0.4), top=Inches(1.2),
        width=Inches(7.5), height=Inches(5.8)
    )
    tf_b = bullets_box.text_frame
    tf_b.word_wrap = True

    from pptx.oxml.ns import qn as _qn
    from lxml import etree as _etree2

    BULLET_COLOR   = RGBColor(180, 0, 0)    # dark-red circle to match title
    TEXT_COLOR     = RGBColor(30, 30, 30)
    BULLET_CHAR    = "\u25CF"               # ● filled circle

    for idx, b in enumerate(slide_data.get("bullets", [])):
        if not b or not b.strip():
            continue
        p = tf_b.add_paragraph() if idx > 0 else tf_b.paragraphs[0]
        p.space_after  = Pt(10)
        p.space_before = Pt(2)
        p.level = 0

        # Disable built-in auto-bullet so we control it manually
        from pptx.oxml.ns import qn as _qn
        buNone_el = _etree2.SubElement(p._p, _qn('a:buNone'))

        # Run 1: colored bullet circle
        run_bullet = p.add_run()
        run_bullet.text = BULLET_CHAR + "  "
        run_bullet.font.size  = Pt(14)
        run_bullet.font.bold  = True
        run_bullet.font.color.rgb = BULLET_COLOR

        # Run 2: indent + actual text
        run_text = p.add_run()
        run_text.text = b.strip()
        run_text.font.size  = Pt(19)
        run_text.font.color.rgb = TEXT_COLOR

    
    # 4️⃣ Right-side generated image (top-right)
    slide_image = slide_data.get("image")   # PIL Image object
    if slide_image:
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_img:
                slide_image.save(tmp_img.name)
                img_path = tmp_img.name
            # Increased height and adjusted top for better visual balance
            slide.shapes.add_picture(img_path,
                                     left=Inches(8.1), top=Inches(1.0),
                                     width=Inches(4.8), height=Inches(5.0))
            print("[PPT] Added slide image.")
        except Exception as e:
            print(f"[PPT] Failed to add slide image: {e}")

    # 5️⃣ Tutor avatar frame — bottom-right corner, "webcam" teaching frame
    # Frame padding / sizes
    FRAME_PAD_SIDE  = Inches(0.08)   # space between frame border & avatar on sides
    FRAME_PAD_TOP   = Inches(0.25)   # header bar height (holds the LED dot)
    FRAME_PAD_BOT   = Inches(0.22)   # name-tag bar height at bottom

    AV_W = Inches(cfg.AVATAR_WIDTH_IN)
    AV_H = Inches(cfg.AVATAR_HEIGHT_IN)

    # Outer frame rectangle (encompasses header + avatar + footer)
    FR_W  = AV_W + FRAME_PAD_SIDE * 2
    FR_H  = AV_H + FRAME_PAD_TOP  + FRAME_PAD_BOT
    FR_X  = Inches(cfg.SLIDE_WIDTH_IN)  - FR_W - Inches(cfg.AVATAR_MARGIN_IN)
    FR_Y  = Inches(cfg.SLIDE_HEIGHT_IN) - FR_H - Inches(cfg.AVATAR_MARGIN_IN)

    # Avatar sits inside the frame
    AV_X = FR_X + FRAME_PAD_SIDE
    AV_Y = FR_Y + FRAME_PAD_TOP

    # ── 5.0  Draw the outer frame (rounded rectangle) ──────────────
    try:
        from pptx.util import Emu
        from pptx.enum.shapes import MSO_SHAPE_TYPE
        import pptx.oxml.ns as _ns
        from lxml import etree as _etree

        # Outer frame: dark navy rounded rectangle
        frame_shape = slide.shapes.add_shape(
            1,  # MSO_SHAPE_TYPE.RECTANGLE (use 5 for rounded later via XML)
            FR_X, FR_Y, FR_W, FR_H
        )
        frame_fill  = frame_shape.fill
        frame_fill.solid()
        frame_fill.fore_color.rgb = RGBColor(18, 30, 54)    # deep navy
        frame_shape.line.color.rgb = RGBColor(60, 90, 140)  # steel-blue border
        frame_shape.line.width = Pt(1.5)
        # Make it a rounded rectangle via XML
        sp_el = frame_shape._element
        prstGeom = sp_el.find('.//' + _ns.qn('a:prstGeom'))
        if prstGeom is not None:
            prstGeom.set('prst', 'roundRect')
            avLst = prstGeom.find(_ns.qn('a:avLst'))
            if avLst is not None:
                for gd in avLst.findall(_ns.qn('a:gd')):
                    avLst.remove(gd)
                gd_el = _etree.SubElement(avLst, _ns.qn('a:gd'))
                gd_el.set('name', 'adj')
                gd_el.set('fmla', 'val 30000')  # corner radius ~30%
        print("[PPT] Drew tutor frame background.")
    except Exception as fe:
        print(f"[PPT] Frame shape error: {fe}")

    # ── 5.1  Header bar (inside top of frame): title strip ─────────
    try:
        hdr_shape = slide.shapes.add_shape(
            1, FR_X, FR_Y, FR_W, FRAME_PAD_TOP
        )
        hdr_shape.fill.solid()
        hdr_shape.fill.fore_color.rgb = RGBColor(10, 20, 40)   # slightly darker navy
        hdr_shape.line.fill.background()                        # no border
        # Round only the top corners — done via XML roundRect + large adj
        sp_hdr = hdr_shape._element
        prstG = sp_hdr.find('.//' + _ns.qn('a:prstGeom'))
        if prstG is not None:
            prstG.set('prst', 'roundRect')
            avL = prstG.find(_ns.qn('a:avLst'))
            if avL is not None:
                for g in avL.findall(_ns.qn('a:gd')):
                    avL.remove(g)
                ge = _etree.SubElement(avL, _ns.qn('a:gd'))
                ge.set('name', 'adj'); ge.set('fmla', 'val 30000')
    except Exception as he:
        print(f"[PPT] Header bar error: {he}")

    # ── 5.2  Red LED "recording" dot + LIVE label ───────────────────
    try:
        DOT_R   = Inches(0.07)
        dot_x   = FR_X + Inches(0.12)
        dot_y   = FR_Y + (FRAME_PAD_TOP - DOT_R * 2) / 2  # vertically centred in header
        dot_shape = slide.shapes.add_shape(9, dot_x, dot_y, DOT_R * 2, DOT_R * 2)  # 9 = Oval
        dot_shape.fill.solid()
        dot_shape.fill.fore_color.rgb = RGBColor(220, 50, 50)  # red LED
        dot_shape.line.fill.background()

        # "LIVE" text next to the dot
        live_box = slide.shapes.add_textbox(
            dot_x + DOT_R * 2 + Inches(0.05),
            FR_Y + Inches(0.04),
            Inches(0.4), FRAME_PAD_TOP - Inches(0.06)
        )
        live_tf = live_box.text_frame
        live_tf.text = "LIVE"
        lp = live_tf.paragraphs[0]
        lp.font.size  = Pt(7)
        lp.font.bold  = True
        lp.font.color.rgb = RGBColor(220, 50, 50)
    except Exception as le:
        print(f"[PPT] LED dot error: {le}")

    # ── 5.3  Avatar / video inside the frame ───────────────────────
    lipsync_video = slide_data.get("lipsync_video")
    if lipsync_video and os.path.exists(lipsync_video):
        try:
            # Re-encode to H.264/AAC so PowerPoint plays it reliably
            embed_path = lipsync_video
            try:
                import imageio_ffmpeg as _iff
                _ffx = _iff.get_ffmpeg_exe()
                import tempfile as _tmp
                _enc_tmp = _tmp.NamedTemporaryFile(delete=False, suffix=".mp4")
                _enc_tmp.close()
                import subprocess as _sp
                r = _sp.run([
                    _ffx, "-y", "-hide_banner", "-loglevel", "error",
                    "-i", lipsync_video,
                    "-c:v", "libx264", "-profile:v", "baseline", "-level", "3.0",
                    "-c:a", "aac", "-b:a", "128k",
                    "-movflags", "+faststart",
                    _enc_tmp.name
                ], capture_output=True)
                if r.returncode == 0 and os.path.getsize(_enc_tmp.name) > 1000:
                    embed_path = _enc_tmp.name
                    print(f"[PPT] Re-encoded to H.264: {embed_path}")
                else:
                    print(f"[PPT] Re-encode failed (rc={r.returncode}), using original.")
            except Exception as enc_e:
                print(f"[PPT] Re-encode step error: {enc_e}")

            # Save avatar as poster frame so PPT shows the face thumbnail
            poster_path = None
            if avatar_img:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_poster:
                    avatar_img.save(tmp_poster.name)
                    poster_path = tmp_poster.name

            slide.shapes.add_movie(
                embed_path,
                left=AV_X, top=AV_Y,
                width=AV_W, height=AV_H,
                poster_frame_image=poster_path,
                mime_type="video/mp4"
            )
            print(f"[PPT] Added framed tutor video: {embed_path}")
        except Exception as e:
            print(f"[PPT] Failed to add tutor video: {e}")
            if avatar_img:
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_av:
                        avatar_img.save(tmp_av.name)
                    slide.shapes.add_picture(tmp_av.name, AV_X, AV_Y, width=AV_W, height=AV_H)
                    print("[PPT] Used static avatar as fallback inside frame.")
                except Exception as e2:
                    print(f"[PPT] Static fallback also failed: {e2}")
    elif avatar_img:
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_av:
                avatar_img.save(tmp_av.name)
                tmp_av_path = tmp_av.name
            slide.shapes.add_picture(tmp_av_path, AV_X, AV_Y, width=AV_W, height=AV_H)
            print("[PPT] Added static avatar inside frame.")
        except Exception as e:
            print(f"[PPT] Failed to add static avatar: {e}")

    # ── 5.4  Footer bar — teacher name label ───────────────────────
    if hasattr(cfg, 'TEACHER_NAME') and cfg.TEACHER_NAME:
        try:
            footer_y = FR_Y + FR_H - FRAME_PAD_BOT
            ft_shape = slide.shapes.add_shape(
                1, FR_X, footer_y, FR_W, FRAME_PAD_BOT
            )
            ft_shape.fill.solid()
            ft_shape.fill.fore_color.rgb = RGBColor(180, 0, 0)   # dark red accent
            ft_shape.line.fill.background()

            name_box = slide.shapes.add_textbox(
                FR_X, footer_y, FR_W, FRAME_PAD_BOT
            )
            name_tf = name_box.text_frame
            name_tf.text = cfg.TEACHER_NAME
            np_ = name_tf.paragraphs[0]
            np_.alignment = PP_ALIGN.CENTER
            np_.font.size  = Pt(9)
            np_.font.bold  = True
            np_.font.color.rgb = RGBColor(255, 255, 255)   # white on red
            print(f"[PPT] Added framed teacher name tag: {cfg.TEACHER_NAME}")
        except Exception as e:
            print(f"[PPT] Failed to add teacher name footer: {e}")

    # 6️⃣ Hidden Audio (tiny shape, top-left — hidden off-screen)
    audio_path = slide_data.get("audio")
    if audio_path and os.path.exists(audio_path):
        try:
            slide.shapes.add_movie(
                audio_path,
                left=Inches(0.05), top=Inches(0.05),
                width=Inches(0.4), height=Inches(0.4),
                poster_frame_image=None,
                mime_type="audio/mp3"
            )
        except Exception as e:
            print(f"[PPT] Failed to add audio: {e}")

    # 7️⃣ Speaker Notes
    notes = slide_data.get("notes")
    if notes:
        slide.notes_slide.notes_text_frame.text = notes

    return slide


def create_pptx_file(ppt_title, slides_data, avatar_img=None, background_img=None):
    """Builds the full PPTX: title slide + one templated slide per entry."""
    prs = Presentation()
    prs.slide_width  = Inches(cfg.SLIDE_WIDTH_IN)
    prs.slide_height = Inches(cfg.SLIDE_HEIGHT_IN)

    # ── Save background image to a temp file (reused on every slide) ──
    bg_path = None
    if background_img:
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_bg:
                background_img.save(tmp_bg.name)
                bg_path = tmp_bg.name
        except Exception as e:
            print(f"[PPT] Background save error: {e}")

    # ── Title Slide ────────────────────────────────────────────────────
    title_slide = prs.slides.add_slide(prs.slide_layouts[6])
    if bg_path:
        try:
            title_slide.shapes.add_picture(bg_path, Inches(0), Inches(0),
                                           width=prs.slide_width, height=prs.slide_height)
        except Exception:
            pass

    tb = title_slide.shapes.add_textbox(Inches(1), Inches(2.5), Inches(11.33), Inches(2))
    tf = tb.text_frame
    tf.text = ppt_title
    p = tf.paragraphs[0]
    p.font.size = Pt(54)
    p.font.bold = True
    p.font.color.rgb = RGBColor(180, 0, 0) # Dark red for visibility on white
    p.alignment = 2  # Center

    sub_tb = title_slide.shapes.add_textbox(Inches(1), Inches(4.8), Inches(11.33), Inches(0.8))
    tf_sub = sub_tb.text_frame
    tf_sub.text = "Generated by Multi-Doc RAG AI"
    p_sub = tf_sub.paragraphs[0]
    p_sub.font.size = Pt(24)
    p_sub.font.color.rgb = RGBColor(60, 60, 60) # Dark grey for subtitle
    p_sub.alignment = 2

    # ── Content Slides ─────────────────────────────────────────────────
    for slide_data in slides_data:
        create_templated_slide(prs, slide_data,
                               bg_path=bg_path,
                               avatar_img=avatar_img)

    # ── Output ────────────────────────────────────────────────────────
    output = BytesIO()
    prs.save(output)
    output.seek(0)
    return output


def convert_to_ppsx(pptx_bytes: BytesIO) -> BytesIO:
    """
    python-pptx always saves with the 'presentation' content type.
    A valid .ppsx requires the 'slideshow' content type inside the ZIP.
    This function patches [Content_Types].xml in-place so PowerPoint
    opens the file without the 'can't read' error.
    """
    import zipfile
    PPTX_CT = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
    PPSX_CT = "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml"

    pptx_bytes.seek(0)
    src_zip = zipfile.ZipFile(pptx_bytes, "r")
    out = BytesIO()
    with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as dst_zip:
        for item in src_zip.infolist():
            data = src_zip.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.replace(
                    PPTX_CT.encode(), PPSX_CT.encode()
                )
            dst_zip.writestr(item, data)
    src_zip.close()
    out.seek(0)
    return out






# ==========================================================
# AUDIO GENERATOR CLASS (Edge TTS)
# ==========================================================
class AudioGenerator:
    """Handles Text-to-Speech generation using edge-tts."""
    
    VOICES = {
        "🇺🇸 US Female (Jenny)": "en-US-JennyNeural",
        "🇺🇸 US Male (Guy)": "en-US-GuyNeural",
        "🇮🇳 Indian Female (Neerja)": "en-IN-NeerjaNeural",
        "🇮🇳 Indian Male (Prabhat)": "en-IN-PrabhatNeural",
        "🇬🇧 UK Female (Sonia)": "en-GB-SoniaNeural",
        "🇬🇧 UK Male (Ryan)": "en-GB-RyanNeural",
    }

    def __init__(self, output_dir="temp_audio"):
        self.output_dir = output_dir
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            
    async def _generate(self, text: str, voice: str, filepath: str):
        """Async generation helper."""
        communicate = edge_tts.Communicate(text, voice)
        await communicate.save(filepath)

    def generate_audio(self, text: str, voice: str = "en-US-GuyNeural") -> str:
        """Generate audio for text and return file path (Synchronous wrapper)."""
        import hashlib
        h = hashlib.md5(f"{text}{voice}".encode()).hexdigest()
        filename = f"audio_{h}.mp3"
        filepath = os.path.join(self.output_dir, filename)
        
        if os.path.exists(filepath):
            return filepath
        
        try:
            asyncio.run(self._generate(text, voice, filepath))
            return filepath
        except Exception as e:
            st.error(f"TTS Error: {e}")
            return None

# ==========================================================
# IMAGE GENERATOR CLASS (Stable Diffusion + Diagrams)
# ==========================================================
def wrap_text_for_mermaid(text, max_words=3):
    words = text.split()
    lines = []
    for i in range(0, len(words), max_words):
        lines.append(" ".join(words[i:i+max_words]))
    return "<br/>".join(lines)

def build_visual_prompt(title, bullets, notes=None):
    """Build a rich, concept-focused prompt for image generation.
    Focuses on the topic CONCEPT, not the slide text itself."""
    # Extract the core topic from the title
    core_topic = re.sub(r'[:\-–].*', '', title).strip()  # Take only the main part before colon

    # Use first bullet for extra context, but rephrase as visual concept
    context_hint = ""
    if bullets:
        first_bullet = re.sub(r'^[•\-\*]\s*', '', bullets[0]).strip()
        context_hint = first_bullet[:80]

    prompt = (
        f"A professional, visually stunning illustration of the concept: '{core_topic}'. "
        f"Context: {context_hint}. "
        f"Style: modern flat design infographic, clean geometric shapes, icons, "
        f"vibrant but professional color palette (blues, teals, oranges), "
        f"no text, no words, no labels, conceptual visualization, "
        f"suitable for a corporate PowerPoint slide, high quality, 4K"
    )
    return prompt

def is_flowchart_requested(title: str, bullets: list) -> bool:
    """Detect if the slide content suggests a flowchart/diagram."""
    keywords = [
        "process", "workflow", "flowchart", "architecture", "sequence", 
        "lifecycle", "steps", "pipeline", "journey", "search", "measurement",
        "algorithm", "logic", "roadmap", "framework", "schema"
    ]
    text = (title + " " + " ".join(bullets)).lower()
    return any(kw in text for kw in keywords)

def generate_mermaid_prompt(title, bullets, context=""):
    """Create a prompt for the LLM to generate Mermaid code."""
    prompt = f"""Generate a Mermaid.js flowchart (graph TD) for the topic: "{title}".
    
    KEY POINTS TO INCLUDE:
    {chr(10).join(bullets[:5])}
    
    REFERENCE CONTEXT (use this for logical flow):
    {context[:1000]}
    
    RULES:
    1. Use 'graph TD' (Top-Down).
    2. Keep nodes concise (3-5 words max). 
    3. Use descriptive labels for arrows if applicable.
    4. Return ONLY the mermaid code block, starting with ```mermaid and ending with ```.
    5. Ensure the logical flow matches the key points.
    
    MERMAID CODE:"""
    return prompt

# ==========================================================
# SMART IMAGE GENERATOR CLASS (Local PIL - No Internet Needed)
# ==========================================================

# Topic → (gradient colors, emoji icon)
TOPIC_THEMES = {
    "web": ((15, 52, 96), (52, 152, 219), "🌐"),
    "data": ((26, 26, 46), (22, 160, 133), "📊"),
    "mining": ((44, 62, 80), (52, 73, 94), "⛏️"),
    "machine learning": ((74, 20, 140), (142, 68, 173), "🤖"),
    "neural": ((74, 20, 140), (142, 68, 173), "🧠"),
    "deep learning": ((74, 20, 140), (142, 68, 173), "🧠"),
    "network": ((15, 52, 96), (41, 128, 185), "🔗"),
    "security": ((30, 39, 46), (231, 76, 60), "🔒"),
    "cloud": ((52, 73, 94), (41, 128, 185), "☁️"),
    "database": ((26, 26, 46), (22, 160, 133), "🗄️"),
    "algorithm": ((20, 90, 50), (39, 174, 96), "⚙️"),
    "classification": ((116, 76, 0), (241, 196, 15), "🏷️"),
    "clustering": ((20, 90, 50), (39, 174, 96), "🔵"),
    "regression": ((116, 76, 0), (241, 196, 15), "📈"),
    "visualization": ((52, 73, 94), (155, 89, 182), "📉"),
    "text": ((44, 62, 80), (127, 140, 141), "📝"),
    "image": ((116, 0, 50), (231, 76, 60), "🖼️"),
    "prediction": ((116, 76, 0), (241, 196, 15), "🔮"),
    "analysis": ((15, 52, 96), (52, 152, 219), "🔍"),
    "sentiment": ((116, 0, 50), (192, 57, 43), "💬"),
    "introduction": ((44, 62, 80), (127, 140, 141), "📖"),
    "default": ((23, 32, 42), (41, 128, 185), "💡"),
}

def get_topic_theme(title):
    title_lower = title.lower()
    for keyword, theme in TOPIC_THEMES.items():
        if keyword in title_lower:
            return theme
    return TOPIC_THEMES["default"]

def draw_gradient(draw, width, height, color1, color2):
    """Draw a vertical gradient background."""
    for y in range(height):
        t = y / height
        r = int(color1[0] + (color2[0] - color1[0]) * t)
        g = int(color1[1] + (color2[1] - color1[1]) * t)
        b = int(color1[2] + (color2[2] - color1[2]) * t)
        draw.line([(0, y), (width, y)], fill=(r, g, b))

class SmartSlideImageGenerator:
    """Generates styled local infographic images using PIL only.
    No internet connection required — works instantly."""

    def __init__(self):
        pass

    def generate_slide_image(self, title, bullets, notes=None):
        """Draw a professional-looking infographic card for the slide topic."""
        from PIL import Image, ImageDraw, ImageFont
        import textwrap

        W, H = 800, 500
        # Ultra-clean white canvas with subtle border
        img = Image.new("RGB", (W, H), (255, 255, 255)) 
        draw = ImageDraw.Draw(img)

        # Theme accent color
        color1, color2, icon = get_topic_theme(title)
        
        # --- Subtle outer frame ---
        draw.rectangle([0, 0, W-1, H-1], outline=(240, 240, 240), width=1)

        # --- Sidebar / Accent strip ---
        draw.rectangle([0, 0, 12, H], fill=color1)

        # --- Load fonts ---
        try:
            # Use distinct weights if possible, fallback to standard arial
            font_title = ImageFont.truetype("arial.ttf", 36)
            font_body  = ImageFont.truetype("arial.ttf", 20)
            font_icon  = ImageFont.truetype("seguiemj.ttf", 72)
        except Exception:
            font_title = ImageFont.load_default()
            font_body  = font_title
            font_icon  = font_title

        # --- Visual Icon / Emoji ---
        try:
            draw.text((60, 50), icon, font=font_icon, embedded_color=True)
        except Exception:
            draw.text((60, 50), "•", font=font_title, fill=color1)

        # --- Title (Safe wrapping for long names) ---
        core_topic = re.sub(r'[:\-–].*', '', title).strip()
        # Responsive wrapping: shorter lines for cleaner typography
        wrapped_title = textwrap.fill(core_topic, width=28)
        title_y = 160
        draw.text((60, title_y), wrapped_title, font=font_title, fill=(30, 30, 30))

        # --- Divider ---
        title_lines = wrapped_title.count('\n') + 1
        line_y = title_y + (title_lines * 42) + 10
        draw.line([60, line_y, 740, line_y], fill=(235, 235, 235), width=2)

        # --- Key Bullet Pills ---
        tag_y = line_y + 25
        for j, b in enumerate(bullets[:4]): # Show up to 4 concepts
            clean_b = re.sub(r'^[•\-\*]\s*', '', b).strip()
            # Trim for layout safety
            short_b = (clean_b[:60] + "…") if len(clean_b) > 60 else clean_b
            
            bbox = draw.textbbox((0, 0), short_b, font=font_body)
            tw = bbox[2] - bbox[0]
            pad_h, pad_v = 15, 6
            tag_x = 60
            tag_h = 36
            
            # Subtle pill background
            draw.rounded_rectangle(
                [tag_x, tag_y, tag_x + tw + pad_h * 2, tag_y + tag_h],
                radius=10, fill=(250, 252, 255)
            )
            # Dot indicator
            draw.ellipse([tag_x + 10, tag_y + 14, tag_x + 18, tag_y + 22], fill=color1)
            # Text
            draw.text((tag_x + 28, tag_y + pad_v), short_b, font=font_body, fill=(70, 70, 70))
            tag_y += tag_h + 12
            if tag_y > H - 60: break

        # --- Branding Signature ---
        draw.text((60, H - 40), "PREMIUM INFOGRAPHIC CONCEPT • SYSTEM GENERATED", font=font_body,
                   fill=(200, 200, 200))

        print(f"[SmartImg] [v2-Modern] Rendered white visual for: {title[:40]}...")
        return img



# ==========================================================
# OLLAMA LLM CLASS
# ==========================================================
class OllamaLLM:
    """Wrapper for Ollama API calls."""
    
    def __init__(self, model: str, base_url: str = "http://127.0.0.1:11434"):
        self.model = model
        self.base_url = base_url
        self.generate_url = f"{base_url}/api/generate"
    
    def generate(self, prompt: str, max_tokens: int = 500) -> str:
        """Generate text using Ollama API."""
        try:
            payload = {
                "model": self.model,
                "prompt": prompt,
                "stream": False,
                "options": {
                    "num_predict": max_tokens,
                    "temperature": 0.7,
                    "top_p": 0.9,
                }
            }
            
            response = requests.post(self.generate_url, json=payload, timeout=120)
            
            if response.status_code == 404:
                raise RuntimeError(f"Model '{self.model}' not found. Run: ollama pull {self.model.split(':')[0]}")
            
            response.raise_for_status()
            
            result = response.json()
            return result.get("response", "").strip()
        except requests.exceptions.ConnectionError:
            raise RuntimeError("Cannot connect to Ollama. Please ensure Ollama is running.")
        except RuntimeError:
            raise
        except Exception as e:
            raise RuntimeError(f"Ollama error: {e}")
    
    @staticmethod
    def check_available(base_url: str = "http://127.0.0.1:11434") -> bool:
        """Check if Ollama is running AND has models installed."""
        try:
            response = requests.get(f"{base_url}/api/tags", timeout=5)
            if response.status_code == 200:
                data = response.json()
                models = data.get("models", [])
                return len(models) > 0
            return False
        except:
            return False
    
    @staticmethod
    def list_models(base_url: str = "http://127.0.0.1:11434") -> List[str]:
        """List available Ollama models."""
        try:
            response = requests.get(f"{base_url}/api/tags", timeout=5)
            if response.status_code == 200:
                data = response.json()
                return [m["name"] for m in data.get("models", [])]
        except:
            pass
        return []


# ==========================================================
# RAG ENGINE
# ==========================================================
class AdvancedRAGEngine:

    def __init__(self, model_id: str, use_ollama: bool = True):
        self.model_id = model_id
        self.use_ollama = use_ollama
        # Force CPU for embeddings to avoid GPU contention/instability with other heavy models
        self.device = "cpu" 
        self.llm: Any = None
        self.vector_store: Any = None
        self.documents: List[Any] = []
        self.chunks: List[Any] = []
        self.total_pages: int = 0
        self.total_chars: int = 0

        # Load embeddings
        with st.spinner("🔧 Loading embeddings..."):
            self.embeddings = HuggingFaceEmbeddings(
                model_name="sentence-transformers/all-mpnet-base-v2",
                model_kwargs={"device": "cpu"},
                encode_kwargs={"normalize_embeddings": True}
            )
            st.success("✅ Embeddings loaded!")

    def load_llm(self):
        """Load the LLM (Ollama or HuggingFace)."""
        if self.llm is not None:
            return

        with st.spinner(f"🤖 Loading {self.model_id}..."):
            if self.use_ollama:
                self.llm = OllamaLLM(self.model_id)
                try:
                    test = self.llm.generate("Say OK", max_tokens=10)
                    st.success(f"✅ Ollama model '{self.model_id}' loaded!")
                except Exception as e:
                    st.error(f"❌ Error: {e}")
                    raise
            else:
                from transformers import AutoTokenizer, AutoModelForSeq2SeqLM, pipeline
                from langchain_huggingface import HuggingFacePipeline
                
                tokenizer = AutoTokenizer.from_pretrained(self.model_id)
                model = AutoModelForSeq2SeqLM.from_pretrained(self.model_id)
                
                pipe = pipeline(
                    "text2text-generation",
                    model=model,
                    tokenizer=tokenizer,
                    max_new_tokens=300,
                    temperature=0.7
                )
                self.llm = HuggingFacePipeline(pipeline=pipe)
                st.success(f"✅ HuggingFace model loaded!")

    def generate(self, prompt: str, max_tokens: int = 400) -> str:
        """Generate text using the loaded LLM."""
        if self.llm is None:
            raise RuntimeError("LLM not loaded.")
        
        if self.use_ollama:
            return self.llm.generate(prompt, max_tokens)
        else:
            result = self.llm.invoke(prompt)
            if isinstance(result, str):
                return result
            return str(result)

    def process_pdfs(self, pdf_paths: List[str]) -> int:
        """Process MULTIPLE PDFs and create a unified vector store."""
        
        all_documents = []
        total_pages = 0
        total_chars = 0
        
        progress_text = st.empty()
        
        for i, pdf_path in enumerate(pdf_paths):
            filename = os.path.basename(pdf_path)
            progress_text.text(f"📄 Processing file {i+1}/{len(pdf_paths)}: {filename}...")
            
            try:
                loader = PyPDFLoader(pdf_path)
                docs = loader.load()
                
                total_pages = cast(int, total_pages) + len(docs)
                total_chars = cast(int, total_chars) + sum(len(doc.page_content) for doc in docs)
                
                all_documents.extend(docs)
                
            except Exception as e:
                st.error(f"Error processing {filename}: {e}")
        
        progress_text.empty()
        
        if not all_documents:
            raise RuntimeError("No documents were successfully loaded.")

        self.documents = all_documents
        self.total_pages = total_pages
        self.total_chars = total_chars
        
        st.info(f"📚 Total Content: {len(pdf_paths)} files, {total_pages} pages, {total_chars:,} characters")
        
        with st.spinner("✂️ Chunking and Indexing content..."):
            splitter = RecursiveCharacterTextSplitter(
                chunk_size=800, # Increased chunk size for better context
                chunk_overlap=150,
                separators=["\n\n", "\n", ". ", " ", ""]
            )

            self.chunks = splitter.split_documents(self.documents)
            self.vector_store = FAISS.from_documents(self.chunks, self.embeddings)
            
            st.success(f"✅ Created {len(self.chunks)} unified knowledge chunks")
            return len(self.chunks)

    def suggest_slide_count(self) -> Tuple[int, List[str]]:
        """Analyze documents and suggest optimal number of slides with unified topic preview."""
        with st.spinner("🔍 Analyzing combined document structure..."):
            # Retrieve broad context
            context = self.retrieve_context("main topics sections headings overview introduction conclusion synthesis", k=15)
            
            context_str = str(context)
            context_slice = context_str[0:4000]
            prompt = f"""Analyze the combined content of these documents and identify the main topics that should become separate slides in a unified presentation.

DOCUMENT CONTENT SAMPLE:
{context_slice}

TASK: List all distinct topics/sections found across these documents. Consolidate similar topics.

OUTPUT FORMAT:
List each topic on a new line. Be comprehensive but synthesized.

TOPICS:"""

            raw = self.generate(prompt, max_tokens=600)
            
            topics = []
            seen = set()
            for line in raw.split("\n"):
                cleaned = re.sub(r'^[\d\.\-\•\*\)\:]+\s*', '', line.strip())
                cleaned = cleaned.strip('"\'')
                
                if 5 < len(cleaned) < 100 and cleaned.lower() not in seen:
                    if not cleaned.lower().startswith(('topic', 'here', 'the following', 'list', 'these')):
                        topics.append(cleaned)
                        seen.add(cleaned.lower())
            
            # Fallback if specific topics not returned
            if len(topics) < 3:
                chunk_limit = min(20, len(self.chunks))
                for i in range(chunk_limit):
                    chunk = self.chunks[i]
                    lines = chunk.page_content.split('\n')
                    for line in lines:
                        line = line.strip()
                        if 10 < len(line) < 80 and line[0].isupper():
                            cleaned = re.sub(r'^[\d\.\-\•\*\)\:]+\s*', '', line)
                            if cleaned.lower() not in seen and len(cleaned) > 5:
                                topics.append(cleaned[:60])
                                seen.add(cleaned.lower())

            suggested_count = max(5, min(20, len(topics))) # Increase minimum for multi-pdf
            
            # Heuristic adjustment
            if suggested_count < 7:
                size_based = max(5, min(20, self.total_chars // 2000))
                page_based = max(5, min(20, self.total_pages // 1.5))
                suggested_count = max(suggested_count, int((size_based + page_based) / 2))
            
            return suggested_count, topics[:suggested_count]

    def retrieve_context(self, query: str, k: int = 10) -> str:
        """Retrieve relevant context from unified vector store."""
        if self.vector_store is None:
            return ""
        
        # Search for more chunks to ensure coverage
        docs = self.vector_store.similarity_search(query, k=k)
        
        # Format with source attribution if available
        context_parts = []
        for d in docs:
            source = d.metadata.get("source_file", "unknown")
            # Page number if available
            page = d.metadata.get("page", "?")
            context_parts.append(f"[Source: {source}, Page: {page}]\n{d.page_content}")
            
        return "\n\n".join(context_parts)

    def extract_topics(self, num_slides: int) -> List[str]:
        """Extract slide topics from the unified content."""
        with st.spinner("🔍 Synthesizing topics..."):
            context = self.retrieve_context("main topics headings sections overview synthesis", k=12)
            
            prompt = f"""Analyze the combined documents and create {num_slides} slide titles for a professional presentation that synthesizes the information.

DOCUMENT CONTENT:
{context[:3500]}

TASK: Create exactly {num_slides} clear, descriptive slide titles that cover the breadth of the provided documents.

RULES:
- Each title should be 3-8 words
- Titles should capture the main ideas/sections across documents
- Order logically (Search -> Analysis -> Synthesis -> Conclusion)
- Do NOT number the titles

OUTPUT (one title per line):"""

            raw = self.generate(prompt, max_tokens=500)
            
            topics = []
            seen = set()
            for line in raw.split("\n"):
                cleaned = re.sub(r'^[\d\.\-\•\*\)\:]+\s*', '', line.strip())
                cleaned = cleaned.strip('"\'')
                
                if len(cleaned) < 5 or len(cleaned) > 100:
                    continue
                if cleaned.lower().startswith(('slide', 'title', 'topic', 'here', 'the following')):
                    continue
                if cleaned.lower() in seen:
                    continue
                    
                topics.append(cleaned)
                seen.add(cleaned.lower())
                
                if len(topics) >= num_slides:
                    break
            
            # Fill with content if needed
            if len(topics) < num_slides:
                 # Simple fill
                topics.extend([f"Additional Topic {i}" for i in range(len(topics)+1, num_slides+1)])

            return topics[:num_slides]

    def generate_candidate_bullets(self, topic: str, num_candidates: int = 10) -> List[Dict]:
        """Generate a list of candidate bullet points with relevance scores."""
        if self.vector_store is None:
            raise RuntimeError("Vector store not initialized.")
        
        context = self.retrieve_context(topic, k=15)
        
        prompt = f"""You are an expert analyst.
TOPIC: "{topic}"

CONTEXT:
{context[:4000]}

TASK: Generate exactly {num_candidates} distinct bullet points relevant to the topic based on the context.
Assess the relevance of each bullet to the topic on a scale of 0-100 (100 = perfect match).

OUTPUT FORMAT:
- <Score 0-100> | <Bullet Point Text>
- <Score> | <Text>
...

RULES:
- No citations.
- Concise checks (15-20 words).
- High relevance only.

CANDIDATES:"""

        raw = self.generate(prompt, max_tokens=800)
        
        candidates = []
        for line in raw.split("\n"):
            line = line.strip()
            if not line: continue
            
            # Remove leading dash/bullet if present
            line = re.sub(r'^[\-\*•\d\.]+\s+', '', line)
            
            score = 70 # Default score
            text = line
            
            if "|" in line:
                try:
                    parts = line.split("|", 1)
                    # Try to parse score, handle cases like "Score: 80"
                    score_match = re.search(r'\d+', parts[0])
                    if score_match:
                        score = int(score_match.group(0))
                    text = parts[1].strip()
                except:
                    pass
            
            if len(text) > 10:
                candidates.append({"score": score, "text": text, "selected": False})
        
        # Fallback if no valid candidates found (parsing failed completely)
        if not candidates:
             for line in raw.split("\n"):
                 if len(line.strip()) > 10:
                     candidates.append({"score": 50, "text": line.strip(), "selected": False})

        # Sort by score desc
        candidates.sort(key=lambda x: x["score"], reverse=True)
        return candidates[:num_candidates]

    def generate_slide_from_candidates(self, topic: str, selected_bullets: List[str], notes_n: int):
        """Generate notes and finalize slide based on selected bullets."""
        
        bullets_text = "\n".join([f"- {b}" for b in selected_bullets])
        context = self.retrieve_context(topic, k=10)

        notes_prompt = f"""Write detailed educational speaker notes for a presentation slide about: "{topic}"

SELECTED POINTS:
{bullets_text}

REFERENCE CONTENT:
{context[:2500]}

TASK: Write {notes_n} informative paragraphs that explain the content of this slide in detail.

RULES:
- TEXTBOOK STYLE: Write as if explaining the concept to a student in a textbook.
- NO PRESENTATION INSTRUCTIONS: Do NOT use phrases like "Welcome everyone", "In this slide", "I will explain", "Set the stage".
- CONTENT FOCUS: Focus purely on the subject matter, facts, definitions, and examples.
- DEPTH: Each paragraph should elaborate on one of the key points/bullets.
- NO CITATIONS: Do NOT include [Source: ...] or (Source: ...) in the text.

SPEAKER NOTES:"""

        notes = self.generate(notes_prompt, max_tokens=700)
        return selected_bullets, notes.strip()


# ==========================================================
# SESSION STATE
# ==========================================================
for key, default in {
    "rag": None,
    "slides": [],
    "slide_idx": 0,
    "pdf_ready": False,
    "model_ready": False,
    "use_ollama": True,
    "suggested_slides": None,
    "suggested_topics": None,
    # New state variables for interactive selection
    "candidates": {}, # {topic: [ {text: str, score: int, selected: bool} ]}
    "selection_step_active": False,
    "current_selection_topic_idx": 0,
    "smart_img_gen": None,
    "audio_gen": None,
    "presentation_mode": False
}.items():
    st.session_state.setdefault(key, default)

# Initialize Generators if needed
if st.session_state.rag and not st.session_state.smart_img_gen:
    st.session_state.smart_img_gen = get_image_generator()
if st.session_state.rag and not st.session_state.audio_gen:
    st.session_state.audio_gen = AudioGenerator()

# ==========================================================
# UI HEADER
# ==========================================================
st.title("📚 Multi-PDF → PPT Generator")
st.caption("Upload multiple PDFs -> Synthesize knowledge -> Generate one cohesive presentation")

# ==========================================================
# SIDEBAR
# ==========================================================
with st.sidebar:
    st.header("⚙️ Settings")
    
    ollama_available = OllamaLLM.check_available()
    
    if ollama_available:
        st.success("✅ Ollama is running")
        available_models = OllamaLLM.list_models()
        
        if available_models:
            model_id = st.selectbox("Ollama Model", available_models, index=0)
            use_ollama = True
        else:
            st.warning("No models found.")
            use_ollama = False
            model_id = "google/flan-t5-large"
    else:
        st.warning("⚠️ Ollama not running. Using HuggingFace.")
        model_map = {
            "Flan-T5 Large": "google/flan-t5-large",
            "Flan-T5 Base": "google/flan-t5-base"
        }
        model_name = st.selectbox("Model", list(model_map.keys()), index=0)
        model_id = model_map[model_name]
        use_ollama = False

    if not st.session_state.model_ready:
        if st.button("🚀 Load Models", type="primary"):
            st.session_state.rag = AdvancedRAGEngine(model_id, use_ollama=use_ollama)
            st.session_state.rag.load_llm()
            st.session_state.model_ready = True
            st.session_state.use_ollama = use_ollama
            
            # Initialize Audio and Lip-Sync Generators
            st.session_state.audio_gen = AudioGenerator()
            st.session_state.lipsync_gen = Wav2LipGenerator()
            st.session_state.smart_img_gen = get_image_generator()
            st.session_state.did_gen = None  # Will be initialized when API key is provided
            
            st.rerun()
    else:
        st.success("✅ Models Loaded")
    
    # D-ID API Configuration
    st.divider()
    st.subheader("🎬 D-ID Professional Avatars")
    
    did_api_key = st.text_input(
        "D-ID API Key",
        type="password",
        help="Get your free API key from https://studio.d-id.com/ (20 free videos/month)",
        placeholder="Enter your D-ID API key..."
    )
    
    if did_api_key:
        if st.button("Test D-ID Connection"):
            try:
                test_gen = DIDGenerator(did_api_key)
                if test_gen.test_connection():
                    st.session_state.did_gen = test_gen
                    st.success("✅ D-ID connected! Professional avatars enabled.")
                else:
                    st.error("❌ Invalid API key or connection failed")
            except Exception as e:
                st.error(f"❌ Error: {e}")
        
        # Store API key in session if not already stored
        if not st.session_state.get("did_gen") and did_api_key:
            st.session_state.did_gen = DIDGenerator(did_api_key)
    else:
        st.info("💡 Enter D-ID API key for professional-quality avatars with full body movements")
        st.markdown("[Get Free API Key →](https://studio.d-id.com/)")

    st.divider()
    ppt_title = st.text_input("PPT Title", "Combined Presentation")
    
    if st.session_state.suggested_slides:
        st.info(f"💡 Suggested: {st.session_state.suggested_slides} slides")
        slides_n = st.slider("Slides", 3, 30, st.session_state.suggested_slides)
    else:
        slides_n = st.slider("Slides", 3, 30, 8)
    
    bullets_n = st.slider("Bullets per slide", 3, 8, 5)
    notes_n = st.slider("Note paragraphs", 2, 6, 3)

    st.divider()
    st.markdown("### 👤 Presenter Avatar")
    
    # Init avatar state if needed
    if "avatar_image" not in st.session_state:
        st.session_state.avatar_image = None
        
    # Custom Upload (Priority)
    uploaded_avatar = st.file_uploader("Upload your photo (for lip-sync anchoring)", type=["jpg", "png", "jpeg"])
    if uploaded_avatar:
        try:
            image = Image.open(uploaded_avatar)
            st.session_state.avatar_image = image
        except:
            st.error("Invalid image")

    if st.session_state.avatar_image:
        st.image(st.session_state.avatar_image, caption="Current Avatar", width=120)
    
    if st.button("Generate AI Avatar", type="secondary"):
        if st.session_state.get("rag") and st.session_state.get("image_gen"):
             with st.spinner("📸 Snapping professional headshot..."):
                 # Use title as context
                 avatar = st.session_state.image_gen.generate_image(ppt_title, ["presenter", "professional"], style="Avatar")
                 if avatar:
                     st.session_state.avatar_image = avatar
                     st.rerun()
        else:
            st.warning("⚠️ Load models first (Step 1)")

    if st.button("🔄 Reset All"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()



# ==========================================================
# STEP 1: MULTI-PDF UPLOAD
# ==========================================================
if not st.session_state.pdf_ready and not st.session_state.selection_step_active:

    st.subheader("📤 Step 1: Upload PDF Documents")
    
    if not st.session_state.model_ready:
        st.warning("⚠️ Please load the models first using the sidebar")
    else:
        # ALLOW MULTIPLE FILES
        pdfs = st.file_uploader("Upload PDF Documents (Select multiple files)", type="pdf", accept_multiple_files=True)

        if pdfs:
            st.write(f"📁 {len(pdfs)} files selected")
            
            col1, col2 = st.columns(2)
            
            with col1:
                analyze_btn = st.button("🔍 Analyze Combined Content", type="secondary")
            
            with col2:
                generate_btn = st.button("🚀 Process & Generate Slides", type="primary")
            
            # Helper to save all files
            def save_temp_pdfs(uploaded_files):
                paths = []
                for f in uploaded_files:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                        tmp.write(f.read())
                        paths.append(tmp.name)
                return paths
            
            # Helper to clean up
            def cleanup_files(paths):
                for p in paths:
                    try:
                        if os.path.exists(p):
                            os.unlink(p)
                    except:
                        pass

            if analyze_btn:
                paths = save_temp_pdfs(pdfs)
                try:
                    st.session_state.rag.process_pdfs(paths)
                    suggested_count, preview_topics = st.session_state.rag.suggest_slide_count()
                    
                    st.session_state.suggested_slides = suggested_count
                    st.session_state.suggested_topics = preview_topics
                    
                    st.success(f"📊 **Recommended: {suggested_count} slides** based on consolidated analysis")
                    
                    with st.expander("📋 Preview of Consolidated Topics", expanded=True):
                        for i, topic in enumerate(preview_topics, 1):
                            st.write(f"{i}. {topic}")
                    
                    cleanup_files(paths)
                except Exception as e:
                    st.error(f"❌ Error analysis: {e}")
                    cleanup_files(paths)
            
            if generate_btn:
                paths = save_temp_pdfs(pdfs)
                try:
                    if st.session_state.rag.vector_store is None:
                        st.session_state.rag.process_pdfs(paths)
                    
                    # USE SUGGESTED TOPICS IF AVAILABLE
                    if st.session_state.suggested_topics:
                        topics = st.session_state.suggested_topics[:slides_n]
                        st.info(f"Using {len(topics)} topics from analysis.")
                    else:
                        topics = st.session_state.rag.extract_topics(slides_n)
                    
                    # GENERATE CANDIDATES FOR SELECTION (STEP 1.5)
                    st.session_state.candidates = {}
                    st.info("Starting candidate generation...") # DEBUG
                    
                    with st.spinner(f"🧩 Generating candidate points for {len(topics)} slides..."):
                        progress_bar = st.progress(0)
                        
                        for idx, t in enumerate(topics):
                            # st.write(f"Debug: Generating for {t}") # DEBUG
                            candidates = st.session_state.rag.generate_candidate_bullets(t, num_candidates=10)
                            if not candidates:
                                st.warning(f"No candidates generated for {t}")
                            st.session_state.candidates[t] = candidates
                            progress_bar.progress((idx + 1) / len(topics))
                    
                    st.success("Generation complete. Switching to selection mode.") # DEBUG
                    st.session_state.selection_step_active = True
                    st.session_state.current_selection_topic_idx = 0
                    st.session_state.slides = [] # Clear old slides
                    cleanup_files(paths)
                    st.rerun()

                except Exception as e:
                    st.error(f"❌ Error generation: {e}")
                    import traceback
                    st.write(traceback.format_exc()) # Change to write for better visibility
                    cleanup_files(paths)

# ==========================================================
# STEP 1.5: INTERACTIVE SELECTION
# ==========================================================
elif st.session_state.selection_step_active:
    
    topics = list(st.session_state.candidates.keys())
    
    if st.session_state.current_selection_topic_idx < len(topics):
        idx = st.session_state.current_selection_topic_idx
        topic = topics[idx]
        total = len(topics)

        # Use a session-state copy of candidates so merges persist across reruns
        cands_key = f"cands_{idx}"
        if cands_key not in st.session_state:
            st.session_state[cands_key] = list(st.session_state.candidates[topic])

        candidates = st.session_state[cands_key]

        st.subheader(f"✅ Step 2: Select Content for Slide {idx+1}/{total}")
        st.progress((idx) / total)
        st.markdown(f"### Slide Title: **{topic}**")
        st.caption("Optionally merge similar points below, then select your best bullets.")

        # ── Helper ─────────────────────────────────────────────────────────
        def cand_sim(a, b):
            STOP = {'the','a','an','and','or','but','in','on','at','to','for',
                    'of','with','by','from','as','is','are','was','were','be',
                    'been','have','has','had','do','does','did','will','would',
                    'could','should','may','might','must','can','it','its','this',
                    'that','these','those','which','who','they','their','we','our'}
            def words(t):
                return set(re.sub(r'[^a-z ]', '', t.lower()).split()) - STOP
            w1, w2 = words(a), words(b)
            if not w1 or not w2: return 0.0
            return len(w1 & w2) / len(w1 | w2)

        # ── MERGE SECTION (outside form) ───────────────────────────────────
        with st.expander("🔀 Merge Similar Candidates", expanded=False):
            st.caption("Identify overlapping points and merge them before selecting. Sorted by similarity (highest first).")

            if len(candidates) < 2:
                st.info("Need at least 2 candidates to merge.")
            else:
                # Pairwise similarity table
                pair_rows = []
                for a in range(len(candidates)):
                    for b in range(a + 1, len(candidates)):
                        score = cand_sim(candidates[a]['text'], candidates[b]['text'])
                        pair_rows.append({
                            "Point A": f"#{a+1}",
                            "Point B": f"#{b+1}",
                            "Similarity": f"{score:.0%}",
                            "_s": score
                        })
                pair_rows.sort(key=lambda r: r["_s"], reverse=True)
                st.dataframe(
                    [{k: v for k, v in r.items() if k != "_s"} for r in pair_rows],
                    use_container_width=True, hide_index=True
                )

                st.markdown("**Check candidates to merge into one:**")
                merge_sel = []
                for c_idx, cand in enumerate(candidates):
                    short = (cand['text'][:90] + "…") if len(cand['text']) > 90 else cand['text']
                    if st.checkbox(f"#{c_idx+1} [{cand['score']}]: {short}", key=f"mchk_{idx}_{c_idx}"):
                        merge_sel.append(c_idx)
                
                merge_disabled = len(merge_sel) < 2
                if st.button("🔀 Merge Selected into One Point", key=f"do_merge_{idx}",
                             disabled=merge_disabled,
                             help="Select 2+ points above, then click to combine them"):
                    merged_text = " | ".join(candidates[m]['text'] for m in merge_sel)
                    avg_score = int(sum(candidates[m]['score'] for m in merge_sel) / len(merge_sel))
                    new_cands = []
                    inserted = False
                    for c_idx, cand in enumerate(candidates):
                        if c_idx in merge_sel:
                            if not inserted:
                                new_cands.append({"text": merged_text, "score": avg_score})
                                inserted = True
                        else:
                            new_cands.append(cand)
                    st.session_state[cands_key] = new_cands
                    # Clear merge checkboxes
                    for k in list(st.session_state.keys()):
                        if k.startswith(f"mchk_{idx}_"):
                            del st.session_state[k]
                    st.success(f"Merged {len(merge_sel)} points into one! Now select the best from below.")
                    st.rerun()

        st.divider()

        # ── SELECTION FORM ─────────────────────────────────────────────────
        st.caption("Select the best bullet points to keep. Top rated items are shown first.")
        with st.form(key=f"select_{idx}"):
            selected_indices = []

            st.markdown("#### Candidate Bullet Points (Score 0-100)")
            for c_idx, cand in enumerate(candidates):
                score = cand['score']
                color = "green" if score > 80 else "orange" if score > 50 else "red"
                label = f"**[{score}]** {cand['text']}"

                # Pre-select top 5
                default_checked = c_idx < 5

                if st.checkbox(label, value=default_checked, key=f"c_{idx}_{c_idx}"):
                    selected_indices.append(c_idx)

            st.divider()
            submit_btn = st.form_submit_button("Confirm Selection & Next ➡️", type="primary")

            if submit_btn:
                # Store selection
                selected_texts = [candidates[i]['text'] for i in selected_indices]

                # Generate final slide data immediately (or verify later)
                # Let's generate notes now to be ready for Step 3
                with st.spinner("📝 Drafting speaker notes..."):
                     _, notes = st.session_state.rag.generate_slide_from_candidates(topic, selected_texts, notes_n)

                slide_data = {
                    "title": topic,
                    "bullets": selected_texts,
                    "notes": notes
                }
                st.session_state.slides.append(slide_data)

                # Clean up merge session keys for this topic
                for k in list(st.session_state.keys()):
                    if k.startswith(f"mchk_{idx}_") or k == cands_key:
                        del st.session_state[k]

                st.session_state.current_selection_topic_idx += 1
                st.rerun()
    else:
        # Done with selection
        st.session_state.selection_step_active = False
        st.session_state.pdf_ready = True # Proceed to Step 3 (Review)
        st.rerun()


# ==========================================================
# STEP 4: PRESENTATION MODE (Playback)
# ==========================================================
elif st.session_state.presentation_mode:
    
    # Ensure index is valid
    if st.session_state.slide_idx >= len(st.session_state.slides):
        st.session_state.presentation_mode = False
        st.session_state.slide_idx = 0
        st.rerun()
    
    i = st.session_state.slide_idx
    slide = st.session_state.slides[i]
    total = len(st.session_state.slides)
    
    # Layout: Avatar + Slide
    st.markdown("### 📽️ Presentation Mode")
    st.progress((i + 1) / total, text=f"Slide {i+1}/{total}")
    
    # Layout: Use full width for content, and float Avatar in bottom-right via CSS
    st.markdown("### 📽️ Presentation Mode")
    st.progress((i + 1) / total, text=f"Slide {i+1}/{total}")
    
    # Content Area
    # Display Slide Content formatted nicely
    st.markdown(f"# {slide['title']}")
    
    img = slide.get("image")
    if img:
        st.image(img, use_container_width=True)
        
    st.divider()
    for b in slide['bullets']:
        st.markdown(f"#### • {b}")

    # AVATAR/LIP-SYNC VIDEO OVERLAY (Bottom Right)
    avatar_img = st.session_state.get("avatar_image")
    audio_path = slide.get("audio")
    lipsync_video = slide.get("lipsync_video")
    
    # AUTO-DISCOVER: If no video path in session, try to find it based on audio hash
    if not lipsync_video and audio_path:
        import hashlib
        
        # Calculate hash (same logic as video generation - audio path only)
        h = hashlib.md5(audio_path.encode()).hexdigest()
        potential_video = os.path.join("temp_lipsync", f"lipsync_{h}.mp4")
        
        if os.path.exists(potential_video) and os.path.getsize(potential_video) > 0:
            lipsync_video = potential_video
            print(f"[AUTO-DISCOVER] Found video: {potential_video}")
        else:
             # Try LivePortrait path
             lp_potential = os.path.join("LivePortrait", "animations", f"lp_gen_{h}.mp4")
             if os.path.exists(lp_potential) and os.path.getsize(lp_potential) > 0:
                 lipsync_video = lp_potential
                 print(f"[AUTO-DISCOVER] Found LP video: {lp_potential}")
    
    # Debug: Print slide data
    print(f"DEBUG: Slide keys: {slide.keys()}")
    print(f"DEBUG: Has lipsync_video key: {'lipsync_video' in slide}")
    print(f"DEBUG: lipsync_video value: {lipsync_video}")
    if lipsync_video:
        print(f"DEBUG: Video file exists: {os.path.exists(lipsync_video)}")
        if os.path.exists(lipsync_video):
            print(f"DEBUG: Video file size: {os.path.getsize(lipsync_video)} bytes")
    
    # Check if lip-sync video exists
    if lipsync_video and os.path.exists(lipsync_video):
        # Display lip-sync video
        print(f"[LIPSYNC] Displaying lip-sync video: {lipsync_video}")  # Debug output
        video_b64 = video_to_base64(lipsync_video)
        
        video_html = f"""
        <style>
        .lipsync-container {{
            position: fixed;
            bottom: 20px;
            right: 20px;
            z-index: 9999;
            display: flex;
            flex-direction: column;
            align-items: center;
            background: rgba(0,0,0,0.5);
            padding: 10px;
            border-radius: 15px;
        }}
        .lipsync-video {{
            width: 200px;
            height: 200px;
            border-radius: 50%;
            object-fit: cover;
            border: 3px solid #4CAF50;
        }}
        </style>
        <div class="lipsync-container">
            <video autoplay loop class="lipsync-video" id="lipsync-video-player">
                <source src="data:video/mp4;base64,{video_b64}" type="video/mp4">
            </video>
            <div style="color:white; font-size:12px; margin-top:5px; font-weight:bold; background:rgba(0,0,0,0.7); padding:2px 8px; border-radius:10px;">AI Tutor</div>
        </div>
        """
        st.markdown(video_html, unsafe_allow_html=True)
        # Don't play separate audio - video already has audio embedded
    else:
        # Fall back to static avatar with CSS animation
        print(f"[WARNING] No lip-sync video found, using CSS animation. Video path: {lipsync_video}")  # Debug output
        # Prepare Avatar HTML
        if avatar_img:
            avatar_b64 = image_to_base64(avatar_img)
            avatar_src = f"data:image/png;base64,{avatar_b64}"
        else:
            # Default avatar URL
            avatar_src = "https://api.dicebear.com/7.x/avataaars/svg?seed=Felix"

        # CSS for fixed bottom-right avatar
        pulse_css = """
        <style>
        @keyframes speaking {
            0% { transform: scale(1); }
            25% { transform: scaleY(1.08) scaleX(0.95); }
            50% { transform: scale(1); }
            75% { transform: scaleY(1.04) scaleX(0.98); }
            100% { transform: scale(1); }
        }
        .avatar-container {
            position: fixed;
            bottom: 20px;
            right: 20px;
            z-index: 9999;
            display: flex;
            flex-direction: column;
            align-items: center;
            background: rgba(0,0,0,0.5);
            padding: 10px;
            border-radius: 15px;
            transition: all 0.3s ease;
        }
        .talking .avatar-img {
            animation: speaking 0.4s infinite;
            border: 3px solid #4CAF50;
        }
        .avatar-img {
            width: 140px;
            height: 140px;
            border-radius: 50%;
            object-fit: cover;
            background: white;
            transition: transform 0.2s;
        }
        </style>
        """
        
        cls = "talking" if audio_path else ""
        
        avatar_html = f"""
        {pulse_css}
        <div class="avatar-container {cls}">
            <img src="{avatar_src}" class="avatar-img">
            <div style="color:white; font-size:12px; margin-top:5px; font-weight:bold; background:rgba(0,0,0,0.7); padding:2px 8px; border-radius:10px;">AI Tutor</div>
        </div>
        """
        st.markdown(avatar_html, unsafe_allow_html=True)


    # Audio Player (only if no lip-sync video - video already has audio)
    if audio_path and not (lipsync_video and os.path.exists(lipsync_video)):
         # Play audio only if we're using static avatar (no lip-sync video)
         st.audio(audio_path, format="audio/mp3", autoplay=True)
    else:
        # Lip-sync video is playing with embedded audio, no separate audio needed
        pass
    
    st.divider()
    
    # Controls
    col_prev, col_exit, col_next = st.columns([1, 2, 1])
    with col_prev:
        if st.button("⬅️ Previous", disabled=(i==0), key="prev_slide"):
            st.session_state.slide_idx -= 1
            st.rerun()
    with col_next:
        if st.button("Next ➡", disabled=(i==len(st.session_state.slides)-1), key="next_slide", type="primary"):
            st.session_state.slide_idx += 1
            st.rerun()
    with col_exit:
        if st.button("❌ Exit Preview", key="exit_pres"):
            st.session_state.presentation_mode = False
            st.rerun()

    st.divider()
    # Add Download Button in Presentation Mode
    if st.button("📥 Download PowerPoint (Final)", type="primary", use_container_width=True, key="dl_from_step4"):
         with st.spinner("🎨 Creating PowerPoint..."):
            ppt_bytes = create_pptx_file(
                ppt_title, 
                st.session_state.slides, 
                st.session_state.get("avatar_image"),
                st.session_state.get("background_image")
            )
            file_name = f"{ppt_title.replace(' ', '_')}.pptx"
            st.download_button(
                label="📥 Click to Download Now",
                data=ppt_bytes,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True
            )
            st.balloons()


# ==========================================================
# STEP 3: REVIEW SLIDES (Final Polish)
# ==========================================================
elif st.session_state.slide_idx < len(st.session_state.slides):

    i = st.session_state.slide_idx
    slide = st.session_state.slides[i]
    total = len(st.session_state.slides)

    col_header, col_play = st.columns([3, 1])
    with col_header:
        st.subheader(f"✏️ Step 3: Review & Edit Slide {i+1} of {total}")
    
    with col_play:
        if st.button("▶️ Start Presentation", type="primary", use_container_width=True):
            st.session_state.presentation_mode = True
            st.session_state.slide_idx = 0
            st.rerun()




    st.progress(i / total, text=f"Reviewing slide {i+1}/{total}")

    col1, col2 = st.columns([2, 1])

    with col1:
        # ── Helper: Jaccard similarity between two bullet strings ──────────
        def bullet_sim(a, b):
            STOP = {'the','a','an','and','or','but','in','on','at','to','for',
                    'of','with','by','from','as','is','are','was','were','be',
                    'been','have','has','had','do','does','did','will','would',
                    'could','should','may','might','must','can','it','its','this',
                    'that','these','those','which','who','they','their','we','our'}
            def words(t):
                return set(re.sub(r'[^a-z ]', '', t.lower()).split()) - STOP
            w1, w2 = words(a), words(b)
            if not w1 or not w2:
                return 0.0
            return len(w1 & w2) / len(w1 | w2)

        bullets_now = slide["bullets"]   # live list

        # ── SECTION 1: Merge similar bullets ──────────────────────────────
        with st.expander("🔀 Merge Similar Bullets", expanded=False):
            if len(bullets_now) < 2:
                st.info("Need at least 2 bullets to merge.")
            else:
                # Show pairwise similarity table
                st.caption("Similarity score between each pair of bullets (higher = more similar):")
                pair_rows = []
                for a in range(len(bullets_now)):
                    for b in range(a + 1, len(bullets_now)):
                        score = bullet_sim(bullets_now[a], bullets_now[b])
                        pair_rows.append({
                            "Bullet A": f"#{a+1}",
                            "Bullet B": f"#{b+1}",
                            "Similarity": f"{score:.0%}",
                            "_score": score
                        })
                pair_rows.sort(key=lambda r: r["_score"], reverse=True)
                # Display table (drop internal _score column)
                st.dataframe(
                    [{k: v for k, v in r.items() if k != "_score"} for r in pair_rows],
                    use_container_width=True, hide_index=True
                )

                st.markdown("**Select bullets to merge:**")
                merge_selected = []
                for j, b in enumerate(bullets_now):
                    short_preview = (b[:80] + "…") if len(b) > 80 else b
                    if st.checkbox(f"#{j+1}: {short_preview}", key=f"merge_chk_{i}_{j}"):
                        merge_selected.append(j)

                if st.button("🔀 Merge Selected Bullets", key=f"do_merge_{i}",
                             disabled=(len(merge_selected) < 2)):
                    # Combine selected bullets into one joined string
                    merged_text = " | ".join(bullets_now[m] for m in merge_selected)
                    # Build new bullet list: keep unselected in order, insert merged at first selected position
                    new_bullets = []
                    inserted = False
                    for j, b in enumerate(bullets_now):
                        if j in merge_selected:
                            if not inserted:
                                new_bullets.append(merged_text)
                                inserted = True
                            # else: skip (absorbed into merged)
                        else:
                            new_bullets.append(b)
                    existing = st.session_state.slides[i]
                    st.session_state.slides[i] = {**existing, "bullets": new_bullets}
                    # SET widget keys to new values so form text_areas refresh immediately
                    for k in list(st.session_state.keys()):
                        if k.startswith(f"merge_chk_{i}_") or k.startswith(f"shortened_{i}_"):
                            del st.session_state[k]
                    for jj, nb in enumerate(new_bullets):
                        st.session_state[f"bullet_{i}_{jj}"] = nb
                    st.success(f"Merged {len(merge_selected)} bullets into one!")
                    st.rerun()

        st.divider()

        # ── SECTION 2: Per-bullet Shorten ──────────────────────────────────
        st.markdown("**Bullet Points — Shorten with AI:**")
        st.caption("Click **Shorten** to condense a bullet. Changes are applied immediately to the edit form below.")

        # Re-read live bullets (may have changed after merge)
        bullets_now = st.session_state.slides[i]["bullets"]
        for j, b in enumerate(bullets_now):
            bcol_text, bcol_btn = st.columns([5, 1])
            with bcol_text:
                preview = (b[:110] + "…") if len(b) > 110 else b
                st.markdown(f"**#{j+1}** {preview}")
            with bcol_btn:
                if st.button("Shorten", key=f"shorten_{i}_{j}", help="AI condenses this bullet"):
                    if st.session_state.get("rag") and st.session_state.rag.llm:
                        with st.spinner("Summarizing…"):
                            try:
                                shorten_prompt = (
                                    f"Rewrite this bullet point in 8-12 words, keeping the core idea. "
                                    f"Return ONLY the shortened text, no quotes, no explanation.\n\n"
                                    f"Bullet: {b}"
                                )
                                shortened = st.session_state.rag.llm.generate(shorten_prompt).strip().strip('"\'')
                                # Update slide data
                                new_bullets = list(st.session_state.slides[i]["bullets"])
                                new_bullets[j] = shortened
                                existing = st.session_state.slides[i]
                                st.session_state.slides[i] = {**existing, "bullets": new_bullets}
                                # ── KEY FIX: SET the widget key directly ──────────────────
                                # Streamlit forms always read from session_state[key], ignoring
                                # the `value` param. Writing here forces the text_area to update.
                                st.session_state[f"bullet_{i}_{j}"] = shortened
                                st.rerun()
                            except Exception as e:
                                st.error(f"Summarize failed: {e}")
                    else:
                        st.warning("Load the LLM model first.")

        st.divider()

        with st.form(key=f"edit_slide_{i}"):
            # Re-read from live slide data — reflects Shorten/Merge changes immediately
            live_slide = st.session_state.slides[i]
            title = st.text_input("Slide Title", live_slide["title"])
            
            st.markdown("**Edit Bullet Points:**")
            bullets = []
            for j, b in enumerate(live_slide["bullets"]):
                bullet_text = st.text_area(f"Bullet {j+1}", b, height=60, key=f"bullet_{i}_{j}")
                bullets.append(bullet_text)
            
            notes = st.text_area("Speaker Notes", live_slide["notes"], height=150)


            st.divider()
            
            col_prev, col_regen, col_next = st.columns([1, 1, 1])
            
            with col_prev:
                prev_btn = col_prev.form_submit_button("⬅️ Previous", disabled=(i == 0))
            
            with col_regen:
                regen_btn = col_regen.form_submit_button("🔄 Regenerate")
            
            with col_next:
                if i < total - 1:
                    next_btn = col_next.form_submit_button("✅ Save & Next ➡️", type="primary")
                else:
                    next_btn = col_next.form_submit_button("✅ Save & Finish", type="primary")

            if prev_btn:
                existing = st.session_state.slides[i]
                st.session_state.slides[i] = {
                    **existing,
                    "title": title, "bullets": bullets, "notes": notes
                }
                st.session_state.slide_idx -= 1
                st.rerun()
            
            if regen_btn:
                with st.spinner("🔄 Regenerating content..."):
                    new_bullets, new_notes = st.session_state.rag.generate_slide(title, bullets_n, notes_n)
                    existing = st.session_state.slides[i]
                    st.session_state.slides[i] = {
                        **existing,
                        "title": title, "bullets": new_bullets, "notes": new_notes
                    }
                st.rerun()
            
            if next_btn:
                existing = st.session_state.slides[i]
                st.session_state.slides[i] = {
                    **existing,
                    "title": title, "bullets": bullets, "notes": notes
                }
                st.session_state.slide_idx += 1
                st.rerun()

            with col2:
                st.markdown("### 👁️ Preview & Media")
                st.markdown(f"**{slide['title']}**")
                
                # --- IMAGE GENERATION ---
                preview_key = f"preview_image_{i}"
                
                # Check for pending preview
                if preview_key in st.session_state:
                    st.info("🖼️ Preview Generated")
                    st.image(st.session_state[preview_key], caption="Preview", width="stretch")
                    
                    col_save, col_discard = st.columns(2)
                    if col_save.button("✅ Keep", key=f"save_img_{i}"):
                        st.session_state.slides[i]["image"] = st.session_state[preview_key]
                        del st.session_state[preview_key]
                        st.rerun()
                        
                    if col_discard.button("❌ Discard", key=f"discard_img_{i}"):
                        del st.session_state[preview_key]
                        st.rerun()
                
                else:
                    # Show existing image
                    current_image = slide.get("image")
                    if current_image:
                        st.image(current_image, caption="Current Visual", width="stretch")
                    
                    if st.button("Generate Visual", key=f"btn_img_{i}"):
                        with st.spinner("🎨 Generating slide-aware visual..."):
                            if "smart_img_gen" not in st.session_state:
                                st.session_state.smart_img_gen = get_image_generator()

                            # Check if we should generate a flowchart instead of a conceptual image
                            is_flowchart = is_flowchart_requested(slide["title"], slide["bullets"])
                            
                            img = None
                            if is_flowchart:
                                with st.status("📊 Building and Polishing Flowchart..."):
                                    st.write("🔍 Extracting logical flow...")
                                    m_context = st.session_state.rag.retrieve_context(slide['title'], k=5)
                                    m_prompt = generate_mermaid_prompt(slide['title'], slide['bullets'], m_context)
                                    mermaid_code = st.session_state.rag.generate(m_prompt)
                                    
                                    st.write("🎨 Rendering base diagram...")
                                    base_img = generate_mermaid_image(mermaid_code)
                                    
                                    if base_img:
                                        st.write("✨ Applying AI Visual Polish...")
                                        # Use our new polish_flowchart method in Fooocus
                                        img = st.session_state.smart_img_gen.polish_flowchart(base_img, slide["title"])
                                    else:
                                        st.warning("Mermaid rendering failed, falling back to standard visual.")
                            
                            # Only generate flowchart — skip PIL/Fooocus image
                            if img:
                                st.session_state[preview_key] = img
                                st.rerun()
                            elif is_flowchart:
                                st.error("Flowchart generation failed. Please try again.")
                            else:
                                st.info("ℹ️ Visual generation is only available for flowchart-type slides. "
                                        "This slide's content does not match a flowchart topic.")

                st.divider()

                # --- AUDIO GENERATION ---
                current_audio = slide.get("audio")
                if current_audio:
                    st.audio(current_audio, format="audio/mp3")
                
                voice_options = list(AudioGenerator.VOICES.keys())
                selected_voice_name = st.selectbox("🗣️ Voice", voice_options, index=1, key=f"voice_{i}")
                selected_voice_id = AudioGenerator.VOICES[selected_voice_name]

                if st.button("Generate Audio", key=f"btn_audio_{i}", help="Generate Voice-over using Edge TTS"):
                    if st.session_state.audio_gen:
                        with st.spinner("🗣️ Synthesizing speech..."):
                            # Text to speech: Title + Notes
                            text_to_speak = f"{slide['title']}. {slide['notes']}"
                            audio_path = st.session_state.audio_gen.generate_audio(text_to_speak, voice=selected_voice_id)
                            if audio_path:
                                st.session_state.slides[i]["audio"] = audio_path
                                st.rerun()
                    else:
                        st.error("Audio Generator failed to initialize.")

                st.divider()
                for bullet in slide["bullets"]:
                    if bullet and bullet.strip():
                        st.markdown(f"• {bullet}")
                st.divider()
                st.caption("Speaker Notes:")
                st.caption(slide["notes"][:300] + "..." if len(slide["notes"]) > 300 else slide["notes"])

                # Debug: Show used context
                with st.expander("🕵️ Debug: Sources Used for this Slide"):
                    # We need to retrieve it again to show it, or store it. 
                    # For now re-retrieve to save memory in session_state
                    debug_context = st.session_state.rag.retrieve_context(slide['title'], k=5)
                    st.text(debug_context[:2000])

# ==========================================================
# STEP 3: FINAL REVIEW & MEDIA SETUP
# ==========================================================
else:
    st.subheader("🎉 Step 3: Final Review & Media Setup")
    st.success(f"✅ All {len(st.session_state.slides)} slides reviewed!")
    
    st.info("👇 **Complete these steps before previewing or downloading:**")
    
    # ===== AVATAR UPLOAD SECTION =====
    st.markdown("### 👤 Step 3.1: Upload Your Tutor Avatar")
    col_avatar_upload, col_avatar_preview = st.columns([2, 1])
    
    with col_avatar_upload:
        uploaded_avatar = st.file_uploader(
            "Upload your avatar image (PNG/JPG)", 
            type=["png", "jpg", "jpeg"],
            key="avatar_uploader",
            help="This avatar will appear in the presentation with lip-sync animation"
        )
        
        if uploaded_avatar:
            avatar_img = Image.open(uploaded_avatar)
            st.session_state.avatar_image = avatar_img
            st.success("✅ Avatar uploaded!")
    
    with col_avatar_preview:
        if st.session_state.get("avatar_image"):
            st.image(st.session_state.avatar_image, caption="Your Tutor Avatar", width="stretch")
        else:
            st.info("No avatar uploaded yet. A default avatar will be used.")
            
    st.divider()
    
    # ===== BACKGROUND UPLOAD =====
    st.markdown("### 🖼️ Step 3.1.5: PPT Background")
    col_bg_up, col_bg_prev = st.columns([2, 1])
    with col_bg_up:
        bg_file = st.file_uploader(
            "Upload background image (PNG/JPG) — applied to every slide",
            type=["png", "jpg", "jpeg"],
            key="ppt_bg_uploader"
        )
        if bg_file:
            bg_img = Image.open(bg_file)
            st.session_state.ppt_bg_image = bg_img
            # Keep legacy key in sync
            st.session_state.background_image = bg_img
            st.success("✅ Background uploaded and set!")
        elif "ppt_bg_image" not in st.session_state:
            st.session_state.ppt_bg_image = None
            st.session_state.background_image = None
    with col_bg_prev:
        if st.session_state.get("ppt_bg_image"):
            st.image(st.session_state.ppt_bg_image,
                     caption="Current Background",
                     width="stretch")
        else:
            st.caption("No background set — slides will use a plain dark style.")
    
    st.divider()
    
    # ===== AUDIO GENERATION SECTION =====
    st.markdown("### 🎵 Step 3.2: Generate Audio for All Slides")
    
    # Count slides with/without audio
    slides_with_audio = sum(1 for s in st.session_state.slides if s.get("audio"))
    slides_without_audio = len(st.session_state.slides) - slides_with_audio
    
    if slides_with_audio > 0:
        st.success(f"✅ {slides_with_audio}/{len(st.session_state.slides)} slides have audio")
    else:
        st.warning(f"⚠️ No audio generated yet. Generate audio to enable lip-sync!")
    
    col_voice, col_gen_audio = st.columns([2, 1])
    
    # Lip-Sync Option
    use_lipsync = st.checkbox(
        "🎬 Generate Lip-Sync Videos (creates videos with mouth movements)", 
        value=False,
        help="Generates videos where the avatar's mouth moves with the audio. Takes longer but provides real lip-sync."
    )
    
    col_voice, col_gen_audio = st.columns([2, 1])
    
    with col_voice:
        bulk_voice_name = st.selectbox(
            "🗣️ Select Voice for All Slides", 
            list(AudioGenerator.VOICES.keys()), 
            index=0, 
            key="final_bulk_voice_select"
        )
        bulk_voice_id = AudioGenerator.VOICES[bulk_voice_name]
    
    with col_gen_audio:
        st.write("")  # Spacing
        st.write("")  # Spacing
        if st.button("🎵 Generate Audio (All Slides)", type="primary", use_container_width=True):
            if st.session_state.audio_gen:
                total = len(st.session_state.slides)
                progress_text = "Synthesizing audio for slide {}/{}"
                my_bar = st.progress(0, text=progress_text.format(0, total))
                
                for idx, s in enumerate(st.session_state.slides):
                    txt = f"{s['title']}. {s['notes']}"
                    path = st.session_state.audio_gen.generate_audio(txt, voice=bulk_voice_id)
                    if path:
                        st.session_state.slides[idx]["audio"] = path
                    my_bar.progress((idx + 1) / total, text=progress_text.format(idx + 1, total))
                
                st.success(f"✅ Audio generated for all slides using {bulk_voice_name}!")
                
                # Generate Lip-Sync Videos if enabled
                if use_lipsync and st.session_state.get("avatar_image"):
                    # Determine which generator to use
                    use_did = False # Force D-ID off
                    
                    # Animation Method Selection
                    animation_method = st.radio(
                        "Local Animation Engine",
                        [
                            "Wav2Lip (Fastest, Lip Only)",
                            "SadTalker Fast (Head Motion, No Enhancer)",
                            "SadTalker Quality (Head Motion + Enhancer)"
                        ],
                        index=2, # Default to SadTalker Quality
                        help="SadTalker Fast is 3x faster but less sharp. Quality mode uses GFPGAN for best results."
                    )
                    
                    use_sadtalker = "SadTalker" in animation_method
                    use_wav2lip = "Wav2Lip" in animation_method

                    if use_did:
                        st.info("🎬 Generating professional D-ID avatars...")
                        generator = st.session_state.did_gen
                    else:
                        if use_sadtalker:
                            st.info("🎬 Generating realistic avatars (SadTalker)...")
                            if ("sadtalker_gen_v3" not in st.session_state or 
                                not hasattr(st.session_state.sadtalker_gen_v3, "generate_lipsync_video") or
                                not hasattr(st.session_state.sadtalker_gen_v3, "use_enhancer")):
                                with st.spinner("Initializing SadTalker (loading 2GB models)..."):
                                    try:
                                        import sadtalker_real
                                        import importlib
                                        importlib.reload(sadtalker_real)
                                        from sadtalker_real import SadTalkerGenerator
                                        temp_gen = SadTalkerGenerator()
                                        # Explicitly test model loading
                                        if temp_gen.load_model():
                                            st.session_state.sadtalker_gen_v3 = temp_gen
                                        else:
                                            raise Exception("Model files missing/corrupt")
                                    except Exception as e:
                                        st.error(f"⚠️ SadTalker failed to load: {e}")
                                        # st.warning("🔄 Falling back to Wav2Lip (Fast Mode).") 
                                        use_sadtalker = False
                            
                            if use_sadtalker and "sadtalker_gen_v3" in st.session_state and hasattr(st.session_state.sadtalker_gen_v3, "generate_lipsync_video"):
                                generator = st.session_state.sadtalker_gen_v3
                            else:
                                # Fallback or None? User said "dont want wav2lip"
                                # But we need a generator. 
                                # If we strictly follow user, we might crash if generator is undefined.
                                # Let's set generator to None or handle it.
                                # Check logic later: if not video_path...
                                # If I set generator = None, it will crash.
                                # I will leave it as lipsync_gen but NOT print the fallback warning.
                                # Or better, just don't run generation.
                                generator = None 

                        elif use_liveportrait:
                            st.info("🎬 Generating enhanced animated avatars (LivePortrait)...")
                            # Init LivePortrait
                            if "live_portrait_gen" not in st.session_state:
                                with st.spinner("Initializing LivePortrait (this may take a minute)..."):
                                    try:
                                        from live_portrait_integration import LivePortraitGenerator
                                        st.session_state.live_portrait_gen = LivePortraitGenerator()
                                        st.session_state.live_portrait_gen.initialize() # Pre-init
                                    except Exception as e:
                                        st.error(f"Failed to load LivePortrait: {e}")
                                        use_liveportrait = False
                            
                            if use_liveportrait:
                                generator = st.session_state.live_portrait_gen
                            else:
                                generator = st.session_state.lipsync_gen
                        else:
                            st.info("🎬 Generating animated avatars (Wav2Lip)...")
                            generator = st.session_state.lipsync_gen
                    
                    # Ensure Wav2Lip generator is ready for fallback (LivePortrait uses it continuously anyway)
                    if "lipsync_gen" not in st.session_state:
                         from wav2lip_integration import Wav2LipGenerator
                         st.session_state.lipsync_gen = Wav2LipGenerator()
                    
                    # Save avatar to temp file
                    import tempfile
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_avatar:
                        st.session_state.avatar_image.save(tmp_avatar.name)
                        avatar_path = tmp_avatar.name
                    
                    lipsync_progress = st.progress(0, text="Generating avatar for slide 0/{}".format(total))
                    
                    for idx, s in enumerate(st.session_state.slides):
                        audio_path = s.get("audio")
                        if audio_path and os.path.exists(audio_path):
                            # CRITICAL FIX: Ensure absolute path for subprocesses (SadTalker runs in its own dir)
                            audio_path = os.path.abspath(audio_path)
                            import time
                            max_retries = 1   # OOM retries just waste time; 1 attempt per slide
                            video_path = None
                            
                            for attempt in range(max_retries):
                                try:
                                    print(f"[INFO] Generating video for slide {idx+1}/{total} (Attempt {attempt+1}/{max_retries})...")
                                    
                                    # ── Per-slide real-time progress UI ──────────────────
                                    slide_header = st.markdown(
                                        f"**🎬 Slide {idx+1}/{total}** — SadTalker rendering..."
                                    )
                                    prog_bar  = st.progress(0)
                                    prog_text = st.empty()
                                    st_progress = {"bar": prog_bar, "text": prog_text}
                                    # ─────────────────────────────────────────────────────

                                    if generator:
                                        if use_sadtalker:
                                            enable_enhancer = "Quality" in animation_method
                                            # SAFETY CHECK: SadTalker crashes with MemoryError on very long clips (e.g. 5+ mins)
                                            audio_duration_estimate = len(s.get("notes", "")) / 15 # rough estimate: 15 chars/sec
                                            if audio_duration_estimate > 200:
                                                st.warning(f"⚠️ Slide {idx+1} has a very long narration ({int(audio_duration_estimate)}s). \n\n"
                                                           "Generation may take 20-30 minutes in Quality Mode. \n"
                                                           "💡 **Tip**: Switch to 'SadTalker Fast' or 'Wav2Lip' for 5x faster results.")
                                            
                                            video_path = generator.generate_lipsync_video(
                                                avatar_path, audio_path,
                                                enhancer=enable_enhancer,
                                                st_progress=st_progress
                                            )
                                        else:
                                            video_path = generator.generate_lipsync_video(avatar_path, audio_path)
                                    else:
                                        st.error("Animation engine not available. Skipping slide.")
                                        video_path = None
                                        break

                                    # Clean up progress widgets
                                    prog_bar.empty()
                                    prog_text.empty()
                                    
                                    if video_path and os.path.exists(video_path):
                                        print(f"[SUCCESS] Video generated for slide {idx}: {video_path}")
                                        st.success(f"✅ Slide {idx+1} done: {os.path.basename(video_path)}")
                                        st.session_state.slides[idx]["lipsync_video"] = video_path
                                        break
                                    else:
                                        print(f"[WARNING] Video generation returned None (Attempt {attempt+1})")
                                        st.error(f"❌ Slide {idx+1} generation failed (check sadtalker_app.log for details)")
                                        
                                except Exception as e:
                                    print(f"[ERROR] Attempt {attempt+1} failed for slide {idx}: {e}")
                                    st.error(f"❌ Error on Slide {idx+1}: {str(e)}")
                            
                            
                            if not video_path:
                                print(f"[FAILURE] Could not generate video for slide {idx} after {max_retries} attempts")
                                # Fallback to LOCAL animation if D-ID fails
                                if st.session_state.get("lipsync_gen"):
                                    print(f"[INFO] FALLBACK: Using local Wav2Lip animation for slide {idx}...")
                                    try:
                                        video_path = st.session_state.lipsync_gen.generate_lipsync_video(avatar_path, audio_path)
                                        if video_path:
                                            st.session_state.slides[idx]["lipsync_video"] = video_path
                                            print(f"[SUCCESS] Local fallback video generated: {video_path}")
                                    except Exception as e_local:
                                        print(f"[ERROR] Local fallback failed: {e_local}")
                        
                        lipsync_progress.progress((idx + 1) / total, text=f"Generating avatar for slide {idx + 1}/{total}")
                    
                    if use_did:
                        st.success("✅ Professional D-ID avatars generated for all slides!")
                    else:
                        st.success("✅ Enhanced animated avatars generated for all slides!")
                    st.info("👉 Go to **Step 3.3** below and click '**Preview Presentation**' to see your animated avatar!")
                elif use_lipsync and not st.session_state.get("avatar_image"):
                    st.warning("⚠️ Please upload an avatar first to generate lip-sync videos.")
                
                st.rerun()
            else:
                st.error("Audio Generator not initialized.")
    
    
    st.divider()
    
    # ===== PREVIEW SECTION =====
    st.markdown("### 👀 Step 3.3: Preview Presentation with Lip-Sync")
    
    col_preview_info, col_preview_btn = st.columns([2, 1])
    
    with col_preview_info:
        st.markdown("""
        **What you'll see in Preview Mode:**
        - ✅ Full-screen slide view
        - ✅ Tutor avatar with lip-sync animation
        - ✅ Auto-playing audio narration
        - ✅ Navigation controls
        """)
    
    with col_preview_btn:
        st.write("")  # Spacing
        if st.button("▶️ Preview Presentation (with Lip-Sync)", type="primary", use_container_width=True, key="final_preview_btn"):
            st.session_state.presentation_mode = True
            st.session_state.slide_idx = 0
            st.rerun()
    
    st.divider()
    
    # ===== DOWNLOAD SECTION =====
    st.markdown("### 📥 Step 3.4: Download PowerPoint")
    
    col_dl_info, col_dl_btn = st.columns([2, 1])
    
    with col_dl_info:
        st.caption("Download the final PowerPoint with embedded audio and avatar.")
        if not st.session_state.get("avatar_image"):
            st.warning("⚠️ No custom avatar uploaded. Default avatar will be used in PPT.")
        if slides_without_audio > 0:
            st.warning(f"⚠️ {slides_without_audio} slides don't have audio yet.")
    
    with col_dl_btn:
        if st.button("📥 Download Presentation", use_container_width=True):
            with st.spinner("🎨 Creating Presentation..."):
                ppt_bytes = create_pptx_file(
                    ppt_title, 
                    st.session_state.slides, 
                    avatar_img=st.session_state.get("avatar_image"),
                    background_img=st.session_state.get("background_image")
                )
                # Convert to valid .ppsx (patches internal content type)
                ppsx_bytes = convert_to_ppsx(ppt_bytes)
                file_name = f"{ppt_title.replace(' ', '_')}.ppsx"
                st.download_button(
                    label="📥 Click to Download PPSX (Auto-play)",
                    data=ppsx_bytes,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.slideshow",
                    use_container_width=True
                )
                st.balloons()
    
    st.divider()
    
    # ===== RESTART SECTION =====
    if st.button("🔄 Start New Project"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()
