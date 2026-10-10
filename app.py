import io
import os
import tempfile
import requests
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from duckduckgo_search import DDGS
from PIL import Image, ImageOps, ImageDraw

st.set_page_config(
    page_title="Agentic PPTX Formatter",
    page_icon="🎨",
    layout="wide"
)

st.title("🎨 Agentic PowerPoint Presentation Builder")

# Theme Palette Options
THEMES = {
    "Modern Corporate": {
        "bg": RGBColor(248, 249, 250),
        "primary": RGBColor(15, 23, 42),
        "accent": RGBColor(37, 99, 235),
        "text": RGBColor(51, 65, 85)
    },
    "Dark Mode Tech": {
        "bg": RGBColor(15, 23, 42),
        "primary": RGBColor(248, 250, 252),
        "accent": RGBColor(56, 189, 248),
        "text": RGBColor(203, 213, 225)
    },
    "Forest & Emerald": {
        "bg": RGBColor(244, 247, 246),
        "primary": RGBColor(20, 83, 45),
        "accent": RGBColor(22, 163, 74),
        "text": RGBColor(55, 65, 81)
    }
}

with st.sidebar:
    st.header("🎨 Style & Layout Controls")
    selected_theme_name = st.selectbox("Select Theme Palette", list(THEMES.keys()))
    theme = THEMES[selected_theme_name]
    
    add_cards = st.checkbox("Enable Title Background Cards", value=True)
    round_image_corners = st.checkbox("Round Image Corners", value=True)

default_slides = """# Agentic AI Systems
- Next-generation automation powered by LLM agents
- Autonomous reasoning, execution, and multi-tool usage
- Real-time web retrieval and multi-agent coordination

---

# Key Architectural Components
- Reasoning Engine: Advanced LLMs for planning and execution
- Tool Integration: APIs, search engines, and database access
- Memory Layer: Short-term contextual and long-term vector storage

---

# Deployment & Future Scalability
- Containerized deployment using Streamlit and Docker
- Enterprise security guardrails and rate-limiting
- Continuous monitoring and feedback loops
"""

slide_text = st.text_area("Slide Content (Markdown Format)", value=default_slides, height=300)

def search_and_process_image(query, rounded=True):
    """Searches web imagery and formats the image with rounded corners."""
    try:
        with DDGS() as ddgs:
            results = list(ddgs.images(query, max_results=1))
            if results:
                image_url = results[0]["image"]
                res = requests.get(image_url, timeout=5)
                if res.status_code == 200:
                    img = Image.open(io.BytesIO(res.content)).convert("RGBA")
                    
                    if rounded:
                        # Apply rounded corner mask
                        size = img.size
                        mask = Image.new('L', size, 0)
                        draw = ImageDraw.Draw(mask)
                        draw.rounded_rectangle((0, 0, size[0], size[1]), radius=25, fill=255)
                        img.putalpha(mask)
                    
                    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                    img.save(temp_file.name, format="PNG")
                    temp_file.close()
                    return temp_file.name
    except Exception as e:
        print(f"Image search error: {e}")
    return None

if st.button("✨ Generate Formatted Presentation", type="primary"):
    if not slide_text.strip():
        st.error("Please enter slide content.")
    else:
        with st.spinner("Applying open-source layout styles and fetching assets..."):
            prs = Presentation()
            prs.slide_width = Inches(13.333)
            prs.slide_height = Inches(7.5)
            blank_layout = prs.slide_layouts[6]
            
            raw_slides = slide_text.split("---")
            
            for raw in raw_slides:
                lines = [l.strip() for l in raw.strip().split("\n") if l.strip()]
                if not lines:
                    continue
                
                title_text = lines[0].replace("#", "").strip()
                body_bullets = lines[1:]
                
                slide = prs.slides.add_slide(blank_layout)
                
                # Set Background Color
                background = slide.background
                fill = background.fill
                fill.solid()
                fill.fore_color.rgb = theme["bg"]
                
                # Add Header Card Accent Box
                if add_cards:
                    header_shape = slide.shapes.add_shape(
                        MSO_SHAPE.ROUNDED_RECTANGLE, 
                        Inches(0.6), Inches(0.5), Inches(12.133), Inches(1.1)
                    )
                    header_shape.fill.solid()
                    header_shape.fill.fore_color.rgb = theme["primary"]
                    header_shape.line.color.rgb = theme["primary"]
                    
                    # Title Inside Card
                    tf = header_shape.text_frame
                    tf.word_wrap = True
                    p = tf.paragraphs[0]
                    p.text = title_text
                    p.font.size = Pt(28)
                    p.font.bold = True
                    p.font.color.rgb = RGBColor(255, 255, 255)
                else:
                    # Clean Minimal Title
                    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11.7), Inches(1.0))
                    tf = title_box.text_frame
                    tf.word_wrap = True
                    p = tf.paragraphs[0]
                    p.text = title_text
                    p.font.size = Pt(32)
                    p.font.bold = True
                    p.font.color.rgb = theme["primary"]
                    
                    # Horizontal Divider Line
                    line_shape = slide.shapes.add_shape(
                        MSO_SHAPE.RECTANGLE,
                        Inches(0.8), Inches(1.5), Inches(11.733), Inches(0.04)
                    )
                    line_shape.fill.solid()
                    line_shape.fill.fore_color.rgb = theme["accent"]
                    line_shape.line.fill.background()
                
                # Fetch & Process Image
                img_path = search_and_process_image(title_text, rounded=round_image_corners)
                
                # Multi-Column Layout Calculation
                has_img = img_path and os.path.exists(img_path)
                content_width = Inches(6.8) if has_img else Inches(11.7)
                
                content_box = slide.shapes.add_textbox(Inches(0.8), Inches(2.0), content_width, Inches(4.8))
                ctf = content_box.text_frame
                ctf.word_wrap = True
                
                for i, b in enumerate(body_bullets):
                    clean_bullet = b.lstrip("-*• ").strip()
                    p_body = ctf.paragraphs[0] if i == 0 else ctf.add_paragraph()
                    p_body.text = clean_bullet
                    p_body.level = 0
                    p_body.font.size = Pt(18)
                    p_body.font.color.rgb = theme["text"]
                    p_body.space_after = Pt(16)
                
                # Place Image with Right-Column Alignment
                if has_img:
                    try:
                        slide.shapes.add_picture(
                            img_path,
                            Inches(8.0),
                            Inches(2.0),
                            width=Inches(4.5)
                        )
                        os.unlink(img_path)
                    except Exception:
                        pass
            
            # Export to download buffer
            buf = io.BytesIO()
            prs.save(buf)
            buf.seek(0)
            
            st.success("Styled presentation generated!")
            st.download_button(
                label="📥 Download Styled Presentation (.pptx)",
                data=buf,
                file_name="styled_presentation.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
