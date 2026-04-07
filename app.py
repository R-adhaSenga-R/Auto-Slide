import streamlit as st
import os
import tempfile
import base64
import json
from mistral_client import MistralClient
from ppt_generator import PPTGenerator
import openai
from dotenv import load_dotenv
import io
import docx2txt
import PyPDF2
import pandas as pd
from audio_recorder_streamlit import audio_recorder
import time
import speech_recognition as sr
import requests
from PIL import Image
from io import BytesIO

# Load environment variables
load_dotenv()

# Set OpenAI API key and initialize client
openai_api_key = os.getenv("OPENAI_API_KEY")
AGENTAI_API_KEY = os.getenv("AGENTAI_API_KEY")
if openai_api_key:
    # Initialize with new client API for OpenAI >= 1.0.0
    openai_client = openai.OpenAI(api_key=openai_api_key)

# Set page config
st.set_page_config(
    page_title="AI Presentation Generator",
    page_icon="📊",
    layout="wide"
)

# Initialize session state for storing generated content
if 'ppt_content' not in st.session_state:
    st.session_state.ppt_content = None
if 'download_ready' not in st.session_state:
    st.session_state.download_ready = False
if 'temp_file_path' not in st.session_state:
    st.session_state.temp_file_path = None
if 'speech_text' not in st.session_state:
    st.session_state.speech_text = ""
if 'file_text' not in st.session_state:
    st.session_state.file_text = ""
if 'is_recording' not in st.session_state:
    st.session_state.is_recording = False
if 'generated_images' not in st.session_state:
    st.session_state.generated_images = {}
if st.button("Generate Presentation"):
    if not prompt.strip():
        st.error("Please enter a topic.")
    else:
        with st.spinner("Generating Presentation..."):
            client = MistralClient()
            ppt_content = client.generate_content(prompt, detailed=True)

            images_dict = {}

            if include_images:
                st.session_state.generated_images = {}
                
                for section in st.session_state.ppt_content.get("sections", []):
                    section_title = section.get("title", "")
                    
                    # Use 'image_prompt' from Mistral AI directly as prompt for image generation
                    image_prompt = section.get("image_prompt", "")
                    
                    # Generate the image using slide content-based prompt
                    image = generate_image(image_prompt, section_title)
                    
                    if image:
                        temp_img_file = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
                        image.save(temp_img_file.name)
                        st.session_state.generated_images[section_title] = temp_img_file.name

            # if include_images:
            #     for section in ppt_content.get("sections", []):
            #         img_prompt = section.get("image_prompt", f"Image related to {section['title']}")
            #         img = generate_image(img_prompt, section['title'])
            #         if img:
            #             temp_img_file = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
            #             img.save(temp_img_file.name)
            #             images[section['title']] = temp_img_file.name

            ppt_gen = PPTGenerator(theme="modern_blue")
            ppt, slide_count = ppt_gen.generate_from_content(ppt_content, images=images)

            temp_dir = tempfile.mkdtemp()
            file_path = os.path.join(temp_dir, "generated_presentation.pptx")
            ppt_gen.save(file_path)

            st.session_state.temp_file_path = file_path
            st.session_state.download_ready = True


def kgenerate_image(slide_content, section_title):
    try:
        url = "https://api-lr.agent.ai/v1/action/generate_image"
        
        payload = {
            "model": "DALL-E 3",
            "model_style": "photo",
            "model_aspect_ratio": "9:16",
            "prompt": slide_content  # Using slide content instead of section title
        }
        
        headers = {
            "Authorization": f"Bearer {AGENTAI_API_KEY}",
            "Content-Type": "application/json"
        }
        
        with st.spinner(f"Generating image for '{section_title}'..."):
            response = requests.post(url, json=payload, headers=headers)
            
            if response.status_code == 200:
                data = response.json()
                if "metadata" in data and "images" in data["metadata"] and len(data["metadata"]["images"]) > 0:
                    image_url = data["metadata"]["images"][0]["url"]
                    image_response = requests.get(image_url)
                    if image_response.status_code == 200:
                        return Image.open(BytesIO(image_response.content))
                    else:
                        st.warning(f"Failed to download image from {image_url}")
                else:
                    st.warning("Image generation failed: No image URL returned")
            else:
                st.warning(f"Image generation request failed with status code {response.status_code}")
        
        return None
    except Exception as e:
        st.error(f"Error generating image: {str(e)}")
        return None
    
import re

def generate_image(slide_content, section_title):
    try:
        url = "https://api-lr.agent.ai/v1/action/generate_image"
        
        payload = {
            "model": "DALL-E 3",
            "model_style": "photo",
            "model_aspect_ratio": "9:16",
            "prompt": slide_content
        }
        
        headers = {
            "Authorization": f"Bearer {AGENTAI_API_KEY}",
            "Content-Type": "application/json"
        }
        
        with st.spinner(f"Generating image for '{section_title}'..."):
            response = requests.post(url, json=payload, headers=headers)
            
            if response.status_code == 200:
                data = response.json()
                
                # Extract the image URL from the HTML in the "response" field
                html_response = data.get("response", "")
                match = re.search(r'<img src="([^"]+)"', html_response)
                image_url = match.group(1) if match else None
                
                if image_url:
                    image_response = requests.get(image_url)
                    if image_response.status_code == 200:
                        return Image.open(BytesIO(image_response.content))
                    else:
                        st.warning(f"Failed to download image from {image_url}")
                else:
                    st.warning(f"No image URL found in API response. Full response: {data}")
            else:
                st.warning(f"Image generation request failed with status code {response.status_code}: {response.text}")
        
        return None
    except Exception as e:
        st.error(f"Error generating image: {str(e)}")
        return None

    
# Function to transcribe speech using OpenAI Whisper - Updated for OpenAI >= 1.0.0
def transcribe_audio(audio_bytes):
    try:
        # Create a recognizer instance
        recognizer = sr.Recognizer()
        
        # Save audio bytes to a temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix='.wav') as tmp_file:
            tmp_file.write(audio_bytes)
            tmp_file_path = tmp_file.name
        
        # Load the audio file and transcribe it
        with sr.AudioFile(tmp_file_path) as source:
            audio_data = recognizer.record(source)
            
            # Use Google's free speech recognition
            # You can also try other engines like recognizer.recognize_sphinx() which is offline
            text = recognizer.recognize_google(audio_data)
            
        # Clean up the temporary file
        os.unlink(tmp_file_path)
        
        return text
        
    except sr.UnknownValueError:
        return "Speech recognition could not understand the audio"
    except sr.RequestError as e:
        return f"Error with speech recognition service: {e}"
    except Exception as e:
        st.error(f"Error transcribing audio: {str(e)}")
        return f"Error: {str(e)}"

# Function to extract text from uploaded files with improved error handling
def extract_text_from_file(uploaded_file):
    text = ""
    file_extension = os.path.splitext(uploaded_file.name)[1].lower()
    
    try:
        # Handle different file types
        if file_extension == '.txt':
            text = uploaded_file.getvalue().decode('utf-8')
        
        elif file_extension == '.docx':
            try:
                text = docx2txt.process(io.BytesIO(uploaded_file.getvalue()))
            except Exception as e:
                return f"Error processing DOCX file: {str(e)}. Make sure it's a valid Word document."
        
        elif file_extension == '.pdf':
            try:
                pdf_reader = PyPDF2.PdfReader(io.BytesIO(uploaded_file.getvalue()))
                for page_num in range(len(pdf_reader.pages)):
                    text += pdf_reader.pages[page_num].extract_text() + "\n"
                
                # Check if we got any text
                if not text.strip():
                    return "The PDF appears to contain scanned images rather than text. Cannot extract content."
            except Exception as e:
                return f"Error processing PDF file: {str(e)}. Make sure it's a valid PDF document."
        
        elif file_extension in ['.csv', '.xlsx', '.xls']:
            try:
                if file_extension == '.csv':
                    df = pd.read_csv(uploaded_file)
                else:
                    df = pd.read_excel(uploaded_file)
                
                # Check if dataframe is empty
                if df.empty:
                    return "The uploaded file appears to be empty."
                
                # Convert the dataframe to a text summary
                text = f"File summary: {uploaded_file.name}\n\n"
                text += f"Columns: {', '.join(df.columns.tolist())}\n"
                text += f"Rows: {len(df)}\n\n"
                text += "Sample data (first 5 rows):\n"
                text += df.head().to_string() + "\n\n"
                text += "Statistical summary:\n"
                
                # Add basic statistics for numerical columns
                numeric_cols = df.select_dtypes(include=['number']).columns
                if len(numeric_cols) > 0:
                    text += df[numeric_cols].describe().to_string()
            except Exception as e:
                return f"Error processing spreadsheet: {str(e)}. Make sure it's a valid file."
        
        else:
            text = f"Unsupported file type: {file_extension}. Please upload a .txt, .docx, .pdf, .csv, or .xlsx file."
    
    except Exception as e:
        text = f"Error processing file: {str(e)}"
    
    # Truncate very large files to prevent issues
    if len(text) > 10000:
        text = text[:10000] + "\n\n... (content truncated for length)"
    
    return text

# Function to download the generated presentation
def get_download_link(file_path, file_name):
    with open(file_path, "rb") as file:
        contents = file.read()
    b64 = base64.b64encode(contents).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,{b64}" download="{file_name}" class="download-button">Download Presentation</a>'
    return href

# Add some custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem !important;
        color: #0072C6;
    }
    .sub-header {
        font-size: 1.5rem !important;
        margin-bottom: 1rem;
    }
    .download-button {
        display: inline-block;
        padding: 0.5em 1em;
        background-color: #0072C6;
        color: white !important;
        text-decoration: none;
        font-weight: bold;
        border-radius: 4px;
        text-align: center;
        transition: background-color 0.3s;
    }
    .download-button:hover {
        background-color: #005999;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 15px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        border-radius: 4px 4px 0 0;
    }
    .input-section {
        background-color: #f8f9fa;
        padding: 1.5rem;
        border-radius: 10px;
        margin-bottom: 1rem;
    }
    .enhanced-text-area {
        border: 1px solid #BBD6EC;
        border-radius: 5px;
    }
</style>
""", unsafe_allow_html=True)

# Title and description
st.markdown('<p class="main-header">AI Presentation Generator</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Generate professional PowerPoint presentations from your input using AI.</p>', unsafe_allow_html=True)

# Main input container
with st.container():
    # Create three columns for different input methods
    col1, col2 = st.columns([3, 2])
    
    with col1:
        # Primary text input - always required
        st.markdown("### ✏️ Presentation Topic or Description")
        prompt = st.text_area(
            "Enter your presentation topic or detailed description:",
            height=150,
            placeholder="Example: 'The impact of artificial intelligence on healthcare in the next decade, covering current technologies, future trends, and ethical considerations.'"
        )
        
        # Supplementary inputs in tabs
        input_tabs = st.tabs(["Voice Input", "Reference Document"])
        
        # Voice input tab
        with input_tabs[0]:
            st.markdown("##### 🎙️ Add Voice Description")
            
            # Store audio_bytes in session state to handle clearing
            if 'audio_bytes' not in st.session_state:
                st.session_state.audio_bytes = None
            
            # Only record if not already cleared
            if not st.session_state.get('cleared_audio', False):
                audio_bytes = audio_recorder(
                    text="Click to start/stop recording",
                    recording_color="#e8585c", 
                    neutral_color="#0072C6",
                    energy_threshold=(-1.0, 1.0),
                    pause_threshold=300.0,
                    sample_rate=44100
                )
                
                if audio_bytes:
                    st.session_state.audio_bytes = audio_bytes
            
            # Add clear recording button
            if st.button("Clear Recording"):
                st.session_state.speech_text = ""
                st.session_state.audio_bytes = None
                st.session_state.cleared_audio = True
                st.rerun()
            
            # Display audio player and transcribe button only if we have audio
            if st.session_state.audio_bytes:
                st.audio(st.session_state.audio_bytes, format="audio/wav")
                
                # Transcribe button
                if st.button("Transcribe Audio"):
                    with st.spinner("Transcribing..."):
                        transcribed_text = transcribe_audio(st.session_state.audio_bytes)
                        if transcribed_text and not transcribed_text.startswith("Error:") and not transcribed_text.startswith("Speech recognition could not understand"):
                            st.session_state.speech_text = transcribed_text
                            st.success("Transcription complete!")
                        else:
                            st.error(transcribed_text or "Failed to transcribe audio. Please try again.")
            
            # Display transcribed text with better formatting
            if st.session_state.speech_text:
                st.markdown("Transcribed Text:")
                st.markdown(f"""
                <div style='background-color:#f0f2f6;padding:10px;border-radius:5px;'>
                    {st.session_state.speech_text}
                </div>
                """, unsafe_allow_html=True)
            
            # Reset the cleared_audio flag when user starts a new session or refreshes
            if st.session_state.get('cleared_audio', False) and not st.session_state.audio_bytes:
                if st.button("Start New Recording"):
                    st.session_state.cleared_audio = False
                    st.rerun()
        
        # File upload tab
        with input_tabs[1]:
            st.markdown("##### 📄 Add Reference Document")
            
            uploaded_file = st.file_uploader(
                "Upload a document to enhance your presentation",
                type=["txt", "pdf", "docx", "csv", "xlsx", "xls"],
                help="Upload research papers, reports, or data to incorporate into your presentation."
            )
            
            if uploaded_file is not None:
                # Process file with progress indicator
                with st.spinner(f"Processing {uploaded_file.name}..."):
                    extracted_text = extract_text_from_file(uploaded_file)
                    
                    # Check if we got an error message
                    if extracted_text.startswith("Error"):
                        st.error(extracted_text)
                    else:
                        st.session_state.file_text = extracted_text
                        
                        # Show success with file details
                        file_size = len(uploaded_file.getvalue()) / 1024  # Size in KB
                        st.success(f"File '{uploaded_file.name}' ({file_size:.1f} KB) successfully processed")
                        
                        # Show preview with expandable section
                        with st.expander("View extracted content", expanded=False):
                            if len(extracted_text) > 1000:
                                preview = extracted_text[:1000] + "... (content truncated for preview)"
                                st.text_area("File content preview", preview, height=200)
                            else:
                                st.text_area("File content", extracted_text, height=200)
    
    with col2:
        # Presentation options
        st.markdown("### Presentation Options")
        
        detailed = st.checkbox("Generate detailed content", value=True, 
                            help="Creates more comprehensive slides with additional information")
        
        theme = st.selectbox(
            "Select presentation theme:",
            ["modern_blue", "elegant_dark", "vibrant", "minimal"],
            index=0,
            help="Visual style for your presentation"
        )
        
        # Add more options if needed
        st.markdown("### Additional Options")
        num_slides = st.slider("Approximate slide count:", 5, 25, 15,
                            help="Target number of slides (actual may vary based on content)")
        
        # Add image generation option
        include_images = st.checkbox("Generate images for slides", value=True,
                                  help="Automatically generate relevant images for your slides")
        
        # Generate button
        if st.button("Generate Presentation", type="primary"):
            # Check if text prompt is provided
            if not prompt.strip():
                st.error("Please provide a presentation topic or description.")
            else:
                # Build the prompt combining all inputs
                full_prompt = prompt.strip()
                
                # Combine with voice input if available
                if st.session_state.speech_text:
                    full_prompt = f"{full_prompt}\n\nAdditional spoken details: {st.session_state.speech_text}"
                
                # Incorporate file content if available
                if st.session_state.file_text:
                    full_prompt = f"{full_prompt}\n\nReference material: {st.session_state.file_text}"
                
                # Add slide count preference
                full_prompt = f"{full_prompt}\n\nTarget exactly {num_slides} slides total."
                
                # Generate the presentation
                with st.spinner("Creating your presentation..."):
                    try:
                        # Initialize Mistral client
                        client = MistralClient()
                        
                        # Generate content
                        st.session_state.ppt_content = client.generate_content(full_prompt, detailed)
                        
                        if "error" in st.session_state.ppt_content:
                            st.error(f"Error: {st.session_state.ppt_content['error']}")
                        else:
                            # Generate images for each section if option is enabled
                            if include_images:
                                st.session_state.generated_images = {}
                                
                                # Generate images for each section
                                for section in st.session_state.ppt_content.get("sections", []):
                                    section_title = section.get("title", "")
                                    # Create an image prompt based on the section content
                                    image_prompt = f"Create a professional, high-quality image for a presentation about: {section_title}"
                                    
                                    # Generate the image
                                    image = generate_image(image_prompt, section_title)
                                    if image:
                                        # Save image to temporary file
                                        temp_img_file = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
                                        image.save(temp_img_file.name)
                                        st.session_state.generated_images[section_title] = temp_img_file.name
                            
                            # Generate PPT with selected theme
                            ppt_gen = PPTGenerator(theme=theme)
                            ppt, actual_slide_count = ppt_gen.generate_from_content(
                                st.session_state.ppt_content, 
                                st.session_state.generated_images if include_images else {}
                            )
                            
                            # Save to temporary file
                            temp_dir = tempfile.mkdtemp()
                            # Use a safe version of the prompt for the filename
                            safe_name = ''.join(c if c.isalnum() else '_' for c in prompt[:20]).strip('_')
                            if not safe_name:
                                safe_name = "ai_presentation"
                            file_name = f"presentation_{safe_name}.pptx"
                            file_path = os.path.join(temp_dir, file_name)
                            ppt_gen.save(file_path)
                            
                            st.session_state.temp_file_path = file_path
                            st.session_state.download_ready = True
                            st.session_state.actual_slide_count = actual_slide_count
                            
                            # Show presentation ready message
                            target_count = num_slides
                            if actual_slide_count == target_count:
                                st.success(f"✅ Your presentation with {actual_slide_count} slides is ready to download!")
                            else:
                                st.warning(f"✅ Your presentation is ready to download! Note: You requested {target_count} slides, but {actual_slide_count} slides were created to best fit the content.")
                    except Exception as e:
                        st.error(f"An error occurred: {str(e)}")
                        st.error("If this is an API error, please check that your Mistral API key is configured correctly in the .env file.")

if st.session_state.download_ready and st.session_state.temp_file_path:
    # Display download button
    st.markdown(get_download_link(st.session_state.temp_file_path, 
                                  os.path.basename(st.session_state.temp_file_path)), 
                unsafe_allow_html=True)
