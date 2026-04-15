import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches
import io
import time
from PIL import Image
from google import genai
from google.genai import types

# --- Global Cache for Deduplication ---
if 'caption_cache' not in st.session_state:
    st.session_state.caption_cache = {}

# --- Advanced Text Extraction (XPath SmartArt Fix) ---
def get_shape_text(shape):
    """Extracts text using both standard API and deep XPath XML drilling."""
    text_content = []
    
    # 1. Standard Text Frames
    if getattr(shape, "has_text_frame", False):
        for paragraph in shape.text_frame.paragraphs:
            text_content.append(paragraph.text)
            
    # 2. Groups (Recursively check subshapes)
    elif getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.GROUP:
        try:
            for subshape in shape.shapes:
                text_content.append(get_shape_text(subshape))
        except AttributeError:
            pass
            
    # 3. SmartArt & Graphic Frames (Deep XPath Dive)
    elif getattr(shape, "shape_type", None) in [19, 24] or hasattr(shape._element, 'nvGraphicFramePr'):
        try:
            # XPath './/a:t' finds ALL text nodes anywhere inside this shape's XML
            for text_node in shape._element.xpath('.//a:t'):
                if text_node.text and text_node.text.strip():
                    text_content.append(text_node.text.strip())
        except Exception:
            pass
        
    return " ".join(text_content).strip()

def get_slide_text(slide):
    text_runs = []
    for shape in slide.shapes:
        text_runs.append(get_shape_text(shape))
    return " ".join(text_runs).strip()

# --- Universal ADA Injection Helper ---
def set_alt_text(shape, alt_text):
    """Injects alt text by aggressively hunting for the description tag."""
    try:
        for prop in ['nvPicPr', 'nvSpPr', 'nvGraphicFramePr', 'nvGrpSpPr']:
            if hasattr(shape._element, prop):
                getattr(shape._element, prop).cNvPr.attrib['descr'] = alt_text
                return
    except Exception:
        pass

def fix_reading_order(slide):
    shapes = list(slide.shapes)
    sortable_shapes = []
    for shape in shapes:
        try:
            if hasattr(shape, "top") and hasattr(shape, "left"):
                if shape.top is not None and shape.left is not None:
                    sortable_shapes.append(shape)
        except Exception:
            pass
            
    sortable_shapes.sort(key=lambda s: (round(s.top / 100000) * 100000, s.left))
    
    if sortable_shapes:
        parent = sortable_shapes[0]._element.getparent()
        for shape in sortable_shapes:
            parent.remove(shape._element)
            parent.append(shape._element)

# --- AI Generation (With Exponential Backoff) ---
def generate_caption(client, image_bytes, prev_text, curr_text, is_diagram=False, diagram_text=""):
    """Calls Gemini with Retry Logic for 503/429 errors."""
    system_prompt = """
    You are an expert in ADA compliance for engineering courses. 
    Generate concise, pedagogical Alt Text (under 125 chars). 
    Focus strictly on the system mechanics, flow, or logical structure depicted.
    """
    
    if is_diagram:
        user_prompt = f"Describe this diagram/SmartArt based on its extracted text: '{diagram_text}'. Context: {curr_text}"
        contents = [user_prompt]
    else:
        image = Image.open(io.BytesIO(image_bytes))
        user_prompt = f"Analyze this image. Context: {curr_text}. Previous context: {prev_text}."
        contents = [image, user_prompt]

    max_retries = 3
    for attempt in range(max_retries):
        try:
            time.sleep(3) # Standard rate limit buffer
            response = client.models.generate_content(
                model='gemini-2.5-flash',
                contents=contents,
                config=types.GenerateContentConfig(system_instruction=system_prompt, temperature=0.2)
            )
            return response.text.strip()
        except Exception as e:
            err_str = str(e)
            if "503" in err_str or "429" in err_str:
                if attempt < max_retries - 1:
                    time.sleep(5 * (attempt + 1)) # Backoff: wait 5s, then 10s
                    continue
            return f"Error: {err_str}"
            
    return "Error: Model overloaded after multiple retries."

def generate_and_add_title(client, slide, slide_text):
    has_title = False
    for shape in slide.shapes:
        if shape.is_placeholder and shape.placeholder_format.type == 1:
            has_title = True
            if not shape.text.strip():
                has_title = False
            break

    if not has_title and slide_text.strip():
        prompt = f"Create a concise, 3-to-6 word title for a presentation slide containing this text. Output ONLY the title.\n\nText: {slide_text}"
        try:
            response = client.models.generate_content(model='gemini-2.5-flash', contents=prompt)
            title_text = response.text.strip()
            txBox = slide.shapes.add_textbox(Inches(-5), Inches(-5), Inches(1), Inches(1))
            txBox.text = f"[Hidden Title] {title_text}"
        except Exception:
            pass

# --- Main App ---
st.set_page_config(page_title="ADA PPTX Automator Pro", layout="centered")
st.title("♿ ADA Course Material Automator (v6)")
st.markdown("Features aggressive image hunting, deep SmartArt XML extraction, and API error retries.")

api_key = st.text_input("Enter your Gemini API Key:", type="password")

st.markdown("### Select ADA Fixes to Apply:")
do_captions = st.checkbox("Generate Image/SmartArt Captions (Alt Text)", value=True)
do_titles = st.checkbox("Generate Missing Slide Titles", value=True)
do_reading_order = st.checkbox("Fix Reading Order (Top-to-Bottom)", value=True)

uploaded_file = st.file_uploader("Upload PowerPoint File", type=["pptx"])

if uploaded_file and api_key:
    if st.button("Process Presentation"):
        client = genai.Client(api_key=api_key)
        prs = Presentation(uploaded_file)
        
        saved_calls = 0
        api_calls = 0
        prev_text = ""
        
        with st.spinner("Processing slides... (This may take a moment due to API rate limits)"):
            progress_bar = st.progress(0)
            total_slides = len(prs.slides)
            
            for i, slide in enumerate(prs.slides):
                curr_text = get_slide_text(slide) 
                
                if do_titles:
                    generate_and_add_title(client, slide, curr_text)
                
                if do_captions:
                    for shape in slide.shapes:
                        # 1. AGGRESSIVE IMAGE HUNTING (Catches standard pictures AND placeholders)
                        try:
                            if hasattr(shape, "image"):
                                img_bytes = shape.image.blob
                                # Use PowerPoint's native sha1 hash for better deduplication
                                img_hash = shape.image.sha1 
                                
                                if img_hash in st.session_state.caption_cache:
                                    set_alt_text(shape, st.session_state.caption_cache[img_hash])
                                    saved_calls += 1
                                else:
                                    caption = generate_caption(client, img_bytes, prev_text, curr_text)
                                    if not caption.startswith("Error"):
                                        st.session_state.caption_cache[img_hash] = caption
                                        set_alt_text(shape, caption)
                                        api_calls += 1
                                    else:
                                        st.warning(f"Slide {i+1} Issue: {caption}")
                                continue # Move to next shape if we successfully handled an image
                        except Exception:
                            pass
                        
                        # 2. SMART_ART / GRAPHIC FRAMES / GROUPS
                        if getattr(shape, "shape_type", None) in [MSO_SHAPE_TYPE.GROUP, 19, 24] or hasattr(shape._element, 'nvGraphicFramePr'):
                            d_text = get_shape_text(shape)
                            if d_text:
                                caption = generate_caption(client, None, prev_text, curr_text, is_diagram=True, diagram_text=d_text)
                                if not caption.startswith("Error"):
                                    set_alt_text(shape, caption)
                                    api_calls += 1
                                    
                if do_reading_order:
                    fix_reading_order(slide)
                
                prev_text = curr_text
                progress_bar.progress((i + 1) / total_slides)

            output = io.BytesIO()
            prs.save(output)
            output.seek(0)
            
            st.success(f"Finished! API Calls: {api_calls} | Redundant Images Saved: {saved_calls}")
            st.download_button("Download ADA File", output, file_name=f"ADA_Compliant_{uploaded_file.name}")