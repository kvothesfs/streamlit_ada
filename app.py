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

# --- Text Extraction Helpers ---
def get_shape_text(shape):
    """Safely extracts text from individual shapes, Groups, and SmartArt."""
    text_content = []
    # Check if the shape has a text frame
    if getattr(shape, "has_text_frame", False):
        for paragraph in shape.text_frame.paragraphs:
            text_content.append(paragraph.text)
    # Check if the shape is a group or SmartArt containing sub-shapes
    elif getattr(shape, "shape_type", None) in [MSO_SHAPE_TYPE.GROUP, MSO_SHAPE_TYPE.SMART_ART]:
        try:
            for subshape in shape.shapes:
                text_content.append(get_shape_text(subshape))
        except AttributeError:
            pass
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

# --- ADA Injection Helpers ---
def set_alt_text(shape, alt_text):
    """Injects alt text into the shape's underlying XML metadata."""
    try:
        if getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.PICTURE:
            shape._element.nvPicPr.cNvPr.attrib['descr'] = alt_text
        else:
            shape._element.nvSpPr.cNvPr.attrib['descr'] = alt_text
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
            
    # Sort primarily by Y-coordinate, secondarily by X-coordinate
    sortable_shapes.sort(key=lambda s: (round(s.top / 100000) * 100000, s.left))
    
    if sortable_shapes:
        parent = sortable_shapes[0]._element.getparent()
        for shape in sortable_shapes:
            parent.remove(shape._element)
            parent.append(shape._element)

# --- AI Generation ---
def generate_caption(client, image_bytes, prev_text, curr_text, is_diagram=False, diagram_text=""):
    """Calls Gemini API to generate alt text, with built-in rate limiting."""
    time.sleep(4) # Rate limit buffer for the free tier (~15 RPM)
    
    system_prompt = """
    You are an expert in ADA compliance for engineering courses. 
    Generate concise, pedagogical Alt Text (under 125 chars). 
    Focus strictly on the system mechanics, flow, or logical structure depicted.
    """
    
    if is_diagram:
        user_prompt = f"Describe this diagram/SmartArt based on its extracted text: '{diagram_text}'. Slide Context: {curr_text}"
        contents = [user_prompt]
    else:
        image = Image.open(io.BytesIO(image_bytes))
        user_prompt = f"Analyze this image. Slide Context: {curr_text}. Previous context: {prev_text}. Example: If this discusses batch processing Gantt charts or facility layout patterns, describe the specific schedules or material flow paths."
        contents = [image, user_prompt]

    try:
        response = client.models.generate_content(
            model='gemini-1.5-flash',
            contents=contents,
            config=types.GenerateContentConfig(system_instruction=system_prompt, temperature=0.2)
        )
        return response.text.strip()
    except Exception as e:
        if "429" in str(e):
            return "RETRY_NEEDED"
        return f"Error: {str(e)}"

def generate_and_add_title(client, slide, slide_text):
    """Generates a hidden title for screen readers if the slide is missing one."""
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
            response = client.models.generate_content(model='gemini-1.5-flash', contents=prompt)
            title_text = response.text.strip()
            
            txBox = slide.shapes.add_textbox(Inches(-5), Inches(-5), Inches(1), Inches(1))
            txBox.text = f"[Hidden Title] {title_text}"
        except Exception:
            pass

# --- Main App ---
st.set_page_config(page_title="ADA PPTX Automator Pro", layout="centered")
st.title("♿ ADA Course Material Automator (v3)")
st.markdown("Upload a `.pptx` file to generate ADA-compliant Alt Text, fix reading order, and generate missing slide titles. Features image deduplication to save API calls.")

api_key = st.text_input("Enter your Gemini API Key:", type="password")

# --- UI Checkboxes Restored ---
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
                # We now safely use get_slide_text to prevent the AttributeError
                curr_text = get_slide_text(slide) 
                
                # 1. Missing Titles
                if do_titles:
                    generate_and_add_title(client, slide, curr_text)
                    titles_added += 1
                
                # 2. Image and SmartArt Captions
                if do_captions:
                    for shape in slide.shapes:
                        # PICTURES
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
                                    st.warning(f"Rate limit reached on slide {i+1}. Try a smaller file.")
                        
                        # SMART_ART / GROUPS
                        elif getattr(shape, "shape_type", None) in [MSO_SHAPE_TYPE.SMART_ART, MSO_SHAPE_TYPE.GROUP]:
                            d_text = get_shape_text(shape)
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