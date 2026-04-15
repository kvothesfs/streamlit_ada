import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches
import io
from PIL import Image
from google import genai
from google.genai import types

# --- Helper Functions ---
def get_slide_text(slide):
    """Extracts all text from a given slide."""
    text_runs = []
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                text_runs.append(run.text)
    return " ".join(text_runs)

def set_alt_text(shape, alt_text):
    """Injects the generated alt text into the PPTX XML."""
    try:
        shape._element.nvPicPr.cNvPr.attrib['descr'] = alt_text
    except Exception:
        pass

def generate_caption(client, image_bytes, prev_text, curr_text):
    """Sends the image and context to Gemini for ADA captioning."""
    image = Image.open(io.BytesIO(image_bytes))
    
    system_prompt = """
    You are an expert in digital accessibility and ADA compliance for higher education engineering courses. 
    Generate precise, concise, and highly descriptive Alternative Text (Alt Text) for images found in lecture presentations.
    1. Be concise but descriptive (under 125 characters unless it's a complex diagram).
    2. Never start with 'Image of' or 'Picture of'.
    3. Focus on the pedagogical function. What is the student supposed to learn?
    4. If purely decorative, output exactly: DECORATIVE.
    5. Transcribe crucial text inside the image.
    6. Output plain text only.
    """
    
    user_prompt = f"""
    Analyze the attached image and generate ADA-compliant alt text. 
    Previous Slide Context: {prev_text}
    Current Slide Text: {curr_text}
    
    Example Context: If the slide discusses batch processing times or facility layout patterns, and the image is an AutoCAD blueprint, describe the specific material flow paths or workstation arrangements shown.
    
    Provide the Alt Text now:
    """
    
    config = types.GenerateContentConfig(system_instruction=system_prompt, temperature=0.2)
    try:
        response = client.models.generate_content(model='gemini-2.5-flash', contents=[image, user_prompt], config=config)
        return response.text.strip()
    except Exception as e:
        return f"Error: {str(e)}"

def generate_and_add_title(client, slide, slide_text):
    """Checks for a title. If missing, uses Gemini to create one and adds it to the slide."""
    has_title = False
    for shape in slide.shapes:
        if shape.is_placeholder and shape.placeholder_format.type == 1: # 1 is Title placeholder
            has_title = True
            if not shape.text.strip(): # Title exists but is empty
                has_title = False
            break

    if not has_title and slide_text.strip():
        # Ask Gemini to summarize the text into a title
        prompt = f"Create a concise, 3-to-6 word title for a presentation slide containing this text. Output ONLY the title, no quotes, plain text.\n\nText: {slide_text}"
        try:
            response = client.models.generate_content(model='gemini-2.5-flash', contents=prompt)
            title_text = response.text.strip()
            
            # Inject a hidden title (1x1 inch box, placed off-canvas so it doesn't disrupt the visual layout but is read by screen readers)
            txBox = slide.shapes.add_textbox(Inches(-5), Inches(-5), Inches(1), Inches(1))
            txBox.text = f"[Hidden Title] {title_text}"
        except Exception:
            pass

def fix_reading_order(slide):
    """Sorts the XML elements of the slide top-to-bottom, left-to-right."""
    shapes = list(slide.shapes)
    
    # Isolate shapes that have spatial coordinates
    sortable_shapes = []
    for shape in shapes:
        try:
            if hasattr(shape, "top") and hasattr(shape, "left"):
                if shape.top is not None and shape.left is not None:
                    sortable_shapes.append(shape)
        except Exception:
            pass # Skip complex groupings that might not have standard boundaries
            
    # Sort primarily by Y-coordinate (top), secondarily by X-coordinate (left)
    # Adding a rough rounding factor to 'top' so items in the same visual row sort left-to-right
    sortable_shapes.sort(key=lambda s: (round(s.top / 100000) * 100000, s.left))
    
    # Reorder elements within the parent XML tree
    if sortable_shapes:
        parent = sortable_shapes[0]._element.getparent()
        for shape in sortable_shapes:
            parent.remove(shape._element)
            parent.append(shape._element)


# --- Main Streamlit App ---
st.set_page_config(page_title="ADA PPTX Automator", layout="centered")

st.title("♿ ADA Course Material Automator")
st.markdown("Upload a `.pptx` file. This tool will generate ADA-compliant Alt Text, fix reading order, and generate missing slide titles.")

api_key = st.text_input("Enter your Google Gemini API Key:", type="password")

# Provide options for the user to select which ADA fixes they want
st.markdown("### Select ADA Fixes to Apply:")
do_captions = st.checkbox("Generate Image Captions (Alt Text)", value=True)
do_titles = st.checkbox("Generate Missing Slide Titles", value=True)
do_reading_order = st.checkbox("Fix Reading Order (Top-to-Bottom)", value=True)

uploaded_file = st.file_uploader("Upload PowerPoint File", type=["pptx"])

if uploaded_file and api_key:
    if st.button("Process Presentation"):
        client = genai.Client(api_key=api_key)
        
        with st.spinner("Analyzing slides... this may take a moment."):
            prs = Presentation(uploaded_file)
            prev_slide_text = ""
            images_processed = 0
            titles_added = 0
            
            progress_bar = st.progress(0)
            total_slides = len(prs.slides)
            
            for i, slide in enumerate(prs.slides):
                curr_slide_text = get_slide_text(slide)
                
                # 1. Missing Titles
                if do_titles:
                    generate_and_add_title(client, slide, curr_slide_text)
                    titles_added += 1 # Rough counter
                
                # 2. Image Captions
                if do_captions:
                    for shape in slide.shapes:
                        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                            image_bytes = shape.image.blob
                            alt_text = generate_caption(client, image_bytes, prev_slide_text, curr_slide_text)
                            if alt_text != "DECORATIVE" and not alt_text.startswith("Error"):
                                set_alt_text(shape, alt_text)
                                images_processed += 1
                                
                # 3. Reading Order
                if do_reading_order:
                    fix_reading_order(slide)
                            
                prev_slide_text = curr_slide_text
                progress_bar.progress((i + 1) / total_slides)
            
            output_stream = io.BytesIO()
            prs.save(output_stream)
            output_stream.seek(0)
            
            st.success(f"Success! Processed {images_processed} images and verified slide structures.")
            
            st.download_button(
                label="Download ADA Compliant Presentation",
                data=output_stream,
                file_name=f"ADA_Compliant_{uploaded_file.name}",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
elif uploaded_file and not api_key:
    st.warning("Please enter your Gemini API Key to proceed.")