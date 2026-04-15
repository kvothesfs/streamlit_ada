import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches
import io
import hashlib
import time
from PIL import Image
from google import genai
from google.genai import types

# --- Global Cache for Deduplication ---
if 'caption_cache' not in st.session_state:
    st.session_state.caption_cache = {}

# --- Advanced Text Extraction (The SmartArt Fix) ---
def get_shape_text(shape):
    """Safely extracts text, including deep XML parsing for SmartArt & Graphic Frames."""
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
            
    # 3. SmartArt & Graphic Frames (Deep XML Dive)
    # Graphic Frames (19) and SmartArt (24) hide text in DrawingML <a:t> nodes
    elif getattr(shape, "shape_type", None) in [19, 24]:
        xml_text = []
        # Iterate through the raw XML tree looking for text nodes
        for node in shape._element.iter():
            if node.tag.endswith('}t') and node.text: # Matches <a:t>
                xml_text.append(node.text)
        text_content.extend(xml_text)
        
    return " ".join(text_content).strip()

def get_slide_text(slide):
    """Extracts all text from an entire slide to provide context to the AI."""
    text_runs = []
    for shape in slide.shapes:
        text_runs.append(get_shape_text(shape))
    return " ".join(text_runs).strip()

def get_image_hash(image_bytes):
    """Generates a unique ID for an image to prevent redundant API calls."""
    return hashlib.md5(image_bytes).hexdigest()

# --- Universal ADA Injection Helper ---
def set_alt_text(shape, alt_text):
    """Injects alt text into the underlying XML metadata for ANY shape type."""
    try:
        # Pictures
        if hasattr(shape._element, 'nvPicPr'):
            shape._element.nvPicPr.cNvPr.attrib['descr'] = alt_text
        # Standard Shapes
        elif hasattr(shape._element, 'nvSpPr'):
            shape._element.nvSpPr.cNvPr.attrib['descr'] = alt_text
        # SmartArt / Graphic Frames
        elif hasattr(shape._element, 'nvGraphicFramePr'):
            shape._element.nvGraphicFramePr.cNvPr.attrib['descr'] = alt_text
        # Groups
        elif hasattr(shape._element, 'nvGrpSpPr'):
            shape._element.nvGrpSpPr.cNvPr.attrib['descr'] = alt_text
    except Exception:
        pass

def fix_reading_order(slide):
    """Sorts the XML elements of the slide top-to-bottom, left-to-right."""
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

# --- AI Generation ---
def generate_caption(client, image_bytes, prev_text, curr_text, is_diagram=False, diagram_text=""):
    """Calls Gemini API to generate alt text, with built-in rate limiting."""
    time.sleep(4) 
    
    system_prompt = """
    You are an expert in ADA compliance for engineering courses. 
    Generate concise, pedagogical Alt Text (under 125 chars). 
    Focus strictly on the system mechanics, flow, or logical structure depicted.
    """
    
    # Model name fixed to gemini-2.5-flash
    model_name = 'gemini-2.5-flash'
    
    if is_diagram:
        user_prompt = f"Describe this diagram/SmartArt based on its extracted text: '{diagram_text}'. Slide Context: {curr_text}"
        contents = [user_prompt]
    else:
        image = Image.open(io.BytesIO(image_bytes))
        user_prompt = f"Analyze this image. Slide Context: {curr_text}. Previous context: {prev_text}."
        contents = [image, user_prompt]

    try:
        response = client.models.generate_content(
            model=model_name,
            contents=contents,
            config=types.GenerateContentConfig(system_instruction=system_prompt, temperature=0.2)
        )
        return response.text.strip()
    except Exception as e:
        if "429" in str(e):
            return "RETRY_NEEDED"
        return f"Error: {str(e)}"

def generate_and_add_title(client, slide, slide_text):
    has_title = False
    for shape in slide.shapes:
        if shape.is_placeholder and shape.placeholder_format.type == 1:
            has_title = True
            if not shape.text.strip():
                has_title = False
            break

    if not has_title and slide_text.strip():
        prompt = f"Create a concise, 3-to-6 word title for a presentation slide containing this text. Output ONLY the title, no quotes.\n\nText: {slide_text}"
        try:
            response = client.models.generate_content(model='gemini-2.5-flash', contents=prompt)
            title_text = response.text.strip()
            
            txBox = slide.shapes.add_textbox(Inches(-5), Inches(-5), Inches(1), Inches(1))
            txBox.text = f"[Hidden Title] {title_text}"
        except Exception:
            pass

# --- Main App ---
st.set_page_config(page_title="ADA PPTX Automator Pro", layout="centered")
st.title("♿ ADA Course Material Automator (v5)")
st.markdown("Upload a `.pptx` file to generate ADA-compliant Alt Text. Features deep XML extraction for SmartArt and image deduplication.")

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
        titles_added = 0
        prev_text = ""
        
        with st.spinner("Processing slides... (This may take a moment due to API rate limits)"):
            progress_bar = st.progress(0)
            total_slides = len(prs.slides)
            
            for i, slide in enumerate(prs.slides):
                curr_text = get_slide_text(slide) 
                
                # 1. Missing Titles
                if do_titles:
                    generate_and_add_title(client, slide, curr_text)
                    titles_added += 1
                
                # 2. Image and SmartArt Captions
                if do_captions:
                    for shape in slide.shapes:
                        # PICTURES (13)
                        if getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.PICTURE:
                            img_bytes = shape.image.blob
                            img_hash = get_image_hash(img_bytes)
                            
                            if img_hash in st.session_state.caption_cache:
                                set_alt_text(shape, st.session_state.caption_cache[img_hash])
                                saved_calls += 1
                            else:
                                caption = generate_caption(client, img_bytes, prev_text, curr_text)
                                if caption != "RETRY_NEEDED":
                                    st.session_state.caption_cache[img_hash] = caption
                                    set_alt_text(shape, caption)
                                    api_calls += 1
                                else:
                                    st.warning(f"Rate limit reached on slide {i+1}.")
                        
                        # SMART_ART (24), GRAPHIC FRAMES (19), GROUPS (6)
                        elif getattr(shape, "shape_type", None) in [MSO_SHAPE_TYPE.GROUP, 19, 24]:
                            d_text = get_shape_text(shape)
                            # Only caption it if we actually found text inside the diagram
                            if d_text:
                                caption = generate_caption(client, None, prev_text, curr_text, is_diagram=True, diagram_text=d_text)
                                if caption != "RETRY_NEEDED":
                                    set_alt_text(shape, caption)
                                    api_calls += 1
                                    
                # 3. Reading Order
                if do_reading_order:
                    fix_reading_order(slide)
                
                prev_text = curr_text
                progress_bar.progress((i + 1) / total_slides)

            # Export
            output = io.BytesIO()
            prs.save(output)
            output.seek(0)
            
            st.success(f"Finished! API Calls: {api_calls} | Redundant Images Saved: {saved_calls}")
            st.download_button("Download ADA File", output, file_name=f"ADA_Compliant_{uploaded_file.name}")