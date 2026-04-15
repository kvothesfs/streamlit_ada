import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import io
import hashlib
import time
from PIL import Image
from google import genai
from google.genai import types

# --- Global Cache for Deduplication ---
if 'caption_cache' not in st.session_state:
    st.session_state.caption_cache = {}

# --- Helper Functions ---
def get_image_hash(image_bytes):
    """Generates a unique ID for an image to prevent redundant API calls."""
    return hashlib.md5(image_bytes).hexdigest()

def get_shape_text(shape):
    """Extracts text from various shape types, including Groups and SmartArt."""
    text_content = []
    if shape.has_text_frame:
        for paragraph in shape.text_frame.paragraphs:
            text_content.append(paragraph.text)
    elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        for subshape in shape.shapes:
            text_content.append(get_shape_text(subshape))
    return " ".join(text_content).strip()

def set_alt_text(shape, alt_text):
    """Injects alt text into the shape's metadata."""
    try:
        # Works for Pictures
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            shape._element.nvPicPr.cNvPr.attrib['descr'] = alt_text
        # Works for other shapes (SmartArt, Groups, etc.)
        else:
            shape._element.nvSpPr.cNvPr.attrib['descr'] = alt_text
    except Exception:
        pass

def generate_caption(client, image_bytes, prev_text, curr_text, is_diagram=False, diagram_text=""):
    """API Call with integrated Rate Limiting."""
    # Free tier limit is often ~15 RPM. 4 seconds ensures we don't exceed it.
    time.sleep(4) 
    
    system_prompt = "You are an expert in ADA compliance for engineering courses. Generate concise, pedagogical Alt Text (under 125 chars)."
    
    if is_diagram:
        user_prompt = f"Describe this diagram/SmartArt based on its text: '{diagram_text}'. Context: {curr_text}"
        contents = [user_prompt]
    else:
        image = Image.open(io.BytesIO(image_bytes))
        user_prompt = f"Analyze this image. Context: {curr_text}. Previous context: {prev_text}"
        contents = [image, user_prompt]

    try:
        response = client.models.generate_content(
            model='gemini-1.5-flash',
            contents=contents,
            config=types.GenerateContentConfig(system_instruction=system_prompt)
        )
        return response.text.strip()
    except Exception as e:
        if "429" in str(e):
            return "RETRY_NEEDED"
        return f"Error: {str(e)}"

# --- Main App ---
st.set_page_config(page_title="ADA PPTX Automator Pro", layout="centered")
st.title("♿ ADA Course Material Automator (v2)")
st.markdown("Optimized with Image Deduplication & SmartArt Support.")

api_key = st.text_input("Enter your Gemini API Key:", type="password")
uploaded_file = st.file_uploader("Upload PowerPoint File", type=["pptx"])

if uploaded_file and api_key:
    if st.button("Process Presentation"):
        client = genai.Client(api_key=api_key)
        prs = Presentation(uploaded_file)
        
        saved_calls = 0
        api_calls = 0
        prev_text = ""
        
        with st.spinner("Processing... Deduplicating images and analyzing structure."):
            progress_bar = st.progress(0)
            total_slides = len(prs.slides)
            
            for i, slide in enumerate(prs.slides):
                curr_text = get_shape_text(slide) # Gets all slide text
                
                for shape in slide.shapes:
                    # CASE 1: Standard Pictures
                    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
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
                                st.warning(f"Rate limit reached on slide {i+1}. Try a smaller file or wait.")
                    
                    # CASE 2: SmartArt or Groups (Treating as Diagrams)
                    elif shape.shape_type in [MSO_SHAPE_TYPE.SMART_ART, MSO_SHAPE_TYPE.GROUP]:
                        d_text = get_shape_text(shape)
                        if d_text:
                            # Use text-based summary for complex diagrams
                            caption = generate_caption(client, None, prev_text, curr_text, is_diagram=True, diagram_text=d_text)
                            set_alt_text(shape, caption)
                            api_calls += 1
                
                prev_text = curr_text
                progress_bar.progress((i + 1) / total_slides)

            # Export
            output = io.BytesIO()
            prs.save(output)
            output.seek(0)
            
            st.success(f"Finished! API Calls: {api_calls} | Redundant Images Saved: {saved_calls}")
            st.download_button("Download ADA File", output, file_name=f"ADA_{uploaded_file.name}")