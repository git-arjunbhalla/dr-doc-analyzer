import streamlit as st
import os
import docx
from pypdf import PdfReader
from google import genai
from google.genai import types
from dotenv import load_dotenv
import math

# Load environment variables for local development
load_dotenv(override=True)

# ==========================================
# 1. PAGE SETUP & CONFIGURATION
# ==========================================
st.set_page_config(
    page_title="Dr.Doc - Document Analyzer & Chatbot",
    page_icon="🩺",
    layout="wide",
    initial_sidebar_state="collapsed" # Hide sidebar by default
)

# Premium Custom CSS Injection for Glassmorphism & Sleek Dark Mode (No Sidebar)
st.markdown("""
<style>
    /* Custom Google Fonts */
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;800&family=Plus+Jakarta+Sans:wght@300;400;600;700&display=swap');
    
    html, body, [data-testid="stAppViewContainer"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
    }
    
    h1, h2, h3, h4, h5, h6 {
        font-family: 'Outfit', sans-serif;
    }
    
    /* Main Background with base dark color and grid pattern */
    [data-testid="stAppViewContainer"] {
        background-color: #05060b;
        background-image: 
            linear-gradient(rgba(255, 255, 255, 0.01) 1px, transparent 1px),
            linear-gradient(90deg, rgba(255, 255, 255, 0.01) 1px, transparent 1px);
        background-size: 40px 40px;
        background-attachment: fixed;
    }
    
    /* Dynamic pulsing backdrop blobs */
    .glow-blob {
        position: fixed;
        width: 600px;
        height: 600px;
        border-radius: 50%;
        filter: blur(120px);
        z-index: -99;
        opacity: 0.28;
        pointer-events: none;
        animation: pulseGlow 15s infinite alternate ease-in-out;
    }
    
    .blob-1 {
        top: -15%;
        left: -10%;
        background: radial-gradient(circle, #8a2be2 0%, rgba(74, 0, 224, 0.4) 70%, transparent 100%);
    }
    
    .blob-2 {
        bottom: -15%;
        right: -10%;
        background: radial-gradient(circle, #00bfff 0%, rgba(0, 34, 62, 0.4) 70%, transparent 100%);
        animation-delay: -7.5s;
    }
    
    @keyframes pulseGlow {
        0% { transform: scale(1) translate(0px, 0px) rotate(0deg); opacity: 0.2; }
        50% { transform: scale(1.18) translate(50px, -30px) rotate(90deg); opacity: 0.38; }
        100% { transform: scale(0.9) translate(-20px, 30px) rotate(180deg); opacity: 0.25; }
    }
    
    /* Animation keyframes for sliding elements */
    @keyframes fadeInUp {
        from {
            opacity: 0;
            transform: translateY(25px);
        }
        to {
            opacity: 1;
            transform: translateY(0);
        }
    }
    
    /* Hide sidebar button to completely ignore sidebar */
    [data-testid="collapsedControl"] {
        display: none !important;
    }
    
    /* Glassmorphism containers with active animations */
    .glass-card {
        background: rgba(13, 16, 31, 0.4);
        backdrop-filter: blur(30px) saturate(210%);
        -webkit-backdrop-filter: blur(30px) saturate(210%);
        border-radius: 24px;
        border: 1px solid rgba(255, 255, 255, 0.08);
        padding: 35px;
        margin-bottom: 25px;
        box-shadow: 0 16px 48px 0 rgba(0, 0, 0, 0.5), inset 0 1px 1px 0 rgba(255, 255, 255, 0.08);
        animation: fadeInUp 0.7s cubic-bezier(0.16, 1, 0.3, 1) both;
        transition: all 0.5s cubic-bezier(0.16, 1, 0.3, 1);
    }
    
    .glass-card:hover {
        border-color: rgba(168, 85, 247, 0.3);
        box-shadow: 0 20px 60px 0 rgba(138, 43, 226, 0.15), inset 0 1px 1px 0 rgba(255, 255, 255, 0.1);
    }
    
    .glass-card-subtle {
        background: rgba(255, 255, 255, 0.01);
        border: 1px solid rgba(255, 255, 255, 0.03);
        border-radius: 16px;
        padding: 20px;
    }
    
    /* Glowing headers with multi-gradient color scale */
    .glow-header {
        background: linear-gradient(135deg, #c084fc 0%, #a855f7 35%, #ec4899 70%, #3b82f6 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 800;
        letter-spacing: -0.8px;
    }
    
    /* Welcome landing page illustration card */
    .welcome-card {
        background: rgba(13, 16, 31, 0.35);
        backdrop-filter: blur(25px);
        border-radius: 28px;
        border: 1px solid rgba(138, 43, 226, 0.15);
        padding: 55px 35px;
        text-align: center;
        margin: 20px auto;
        max-width: 800px;
        box-shadow: 0 20px 60px rgba(0, 0, 0, 0.6), inset 0 1px 1px 0 rgba(255, 255, 255, 0.05);
        animation: fadeInUp 0.7s cubic-bezier(0.16, 1, 0.3, 1) both;
    }
    
    /* Document statistics badges */
    .stats-badge {
        background: rgba(255, 255, 255, 0.015) !important;
        border: 1px solid rgba(255, 255, 255, 0.06) !important;
        border-radius: 16px;
        padding: 16px;
        text-align: center;
        transition: all 0.4s cubic-bezier(0.16, 1, 0.3, 1);
        box-shadow: inset 0 1px 0 0 rgba(255, 255, 255, 0.05);
    }
    
    .stats-badge:hover {
        background: rgba(138, 43, 226, 0.08) !important;
        border-color: rgba(168, 85, 247, 0.4) !important;
        transform: translateY(-4px) scale(1.03);
        box-shadow: 0 15px 30px rgba(138, 43, 226, 0.2), inset 0 1px 0 0 rgba(255, 255, 255, 0.1);
    }
    
    .stats-label {
        font-size: 11px;
        color: #94a3b8;
        text-transform: uppercase;
        font-weight: 600;
        letter-spacing: 0.5px;
    }
    
    .stats-value {
        font-size: 22px;
        font-weight: 700;
        background: linear-gradient(135deg, #f3e8ff, #c084fc);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-top: 4px;
        font-family: 'Outfit', sans-serif;
    }
    
    /* Custom tabs pill styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 12px;
        background-color: transparent;
        padding-bottom: 12px;
        border-bottom: 1px solid rgba(255, 255, 255, 0.05);
    }
    
    .stTabs [data-baseweb="tab"] {
        background-color: rgba(255, 255, 255, 0.02);
        border: 1px solid rgba(255, 255, 255, 0.05);
        border-radius: 12px;
        padding: 12px 24px;
        color: #94a3b8;
        font-weight: 600;
        font-size: 15px;
        transition: all 0.3s cubic-bezier(0.16, 1, 0.3, 1);
    }
    
    .stTabs [data-baseweb="tab"]:hover {
        color: #c084fc;
        border-color: rgba(168, 85, 247, 0.3);
        background-color: rgba(138, 43, 226, 0.06);
    }
    
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, rgba(138, 43, 226, 0.22) 0%, rgba(79, 70, 229, 0.22) 100%) !important;
        border-color: rgba(168, 85, 247, 0.5) !important;
        color: #ffffff !important;
        box-shadow: 0 0 20px rgba(138, 43, 226, 0.25), inset 0 1px 0 0 rgba(255, 255, 255, 0.1) !important;
    }

    /* Custom suggestions button design */
    .suggestion-btn {
        margin: 5px;
        border-radius: 20px;
        background-color: rgba(138, 43, 226, 0.1);
        border: 1px solid rgba(138, 43, 226, 0.3);
        color: #c084fc;
        padding: 6px 12px;
        font-size: 13px;
        cursor: pointer;
        display: inline-block;
        transition: all 0.2s ease;
    }
    
    .suggestion-btn:hover {
        background-color: rgba(138, 43, 226, 0.25);
        border-color: rgba(138, 43, 226, 0.6);
        color: #e9d5ff;
    }

    /* Code and preform text layout override */
    .raw-text-container {
        max-height: 500px;
        overflow-y: auto;
        padding: 20px;
        background: rgba(5, 6, 11, 0.85);
        border-radius: 16px;
        border: 1px solid rgba(255, 255, 255, 0.05);
        font-family: 'Courier New', Courier, monospace;
        font-size: 14px;
        line-height: 1.6;
        color: #cbd5e1;
        white-space: pre-wrap;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. STATE MANAGEMENT & SESSION KEYS
# ==========================================
if "current_file_name" not in st.session_state:
    st.session_state.current_file_name = None
if "document_text" not in st.session_state:
    st.session_state.document_text = ""
if "document_summary" not in st.session_state:
    st.session_state.document_summary = ""
if "messages" not in st.session_state:
    st.session_state.messages = []
if "chat_session" not in st.session_state:
    st.session_state.chat_session = None

# ==========================================
# 3. HELPER FUNCTIONS
# ==========================================
def extract_text_from_pdf(file):
    """Extracts text from uploaded PDF file."""
    try:
        reader = PdfReader(file)
        text = ""
        for page in reader.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
        return text.strip()
    except Exception as e:
        st.error(f"Error reading PDF: {e}")
        return None

def extract_text_from_docx(file):
    """Extracts text (including tables) from uploaded DOCX file."""
    try:
        doc = docx.Document(file)
        full_text = []
        
        # Paragraphs
        for para in doc.paragraphs:
            if para.text.strip():
                full_text.append(para.text)
                
        # Tables
        for table in doc.tables:
            for row in table.rows:
                row_text = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                # De-duplicate adjacent identical cell values (happens with merged cells)
                clean_row = []
                for val in row_text:
                    if not clean_row or clean_row[-1] != val:
                        clean_row.append(val)
                if clean_row:
                    full_text.append(" | ".join(clean_row))
                    
        return "\n".join(full_text).strip()
    except Exception as e:
        st.error(f"Error reading DOCX: {e}")
        return None

def get_gemini_client(api_key):
    """Initializes and returns a Gemini client."""
    return genai.Client(api_key=api_key)

def generate_summary(text, client):
    """Generates a structured markdown summary of the text."""
    prompt = f"""
    You are an expert document analyst. Analyze the following document text and provide a comprehensive, structured summary in Markdown.
    
    Format requirements:
    # 📝 Document Executive Summary
    
    ## 🔍 Overview & Context
    [Provide a detailed, high-level summary of the document, its core subject, background context, and overall objectives. (approx 2-3 paragraphs)]
    
    ## 🎯 Key Takeaways & Core Themes
    - **[Theme 1]**: [Detail about theme 1]
    - **[Theme 2]**: [Detail about theme 2]
    - **[Theme 3]**: [Detail about theme 3]
    (Provide 3-6 critical takeaways highlighting primary insights, findings, or clauses)
    
    ## ⚡ Action Items, Milestones & Directives
    - [ ] **[Action Item/Requirement 1]**: [Describe who is responsible or what needs to be done]
    - [ ] **[Action Item/Requirement 2]**: [Describe context/milestone]
    (Highlight any explicit tasks, deadlines, action items, or recommendations)
    
    ## 📊 Document Quick Profile
    - **Complexity & Readability**: [Low / Medium / High - explain why briefly]
    - **Target Audience**: [Identify who the target audience/stakeholder is]
    - **Notable Keywords & Concepts**: [List 5-8 crucial terms or concepts referenced in the document]
    
    Document Text:
    ---
    {text}
    ---
    """
    
    try:
        response = client.models.generate_content(
            model="gemini-3.5-flash",
            contents=prompt
        )
        return response.text
    except Exception as e:
        st.error(f"Failed to generate summary: {e}")
        return None

# ==========================================
# 4. CREDENTIALS AUTHENTICATION (AUTOMATIC)
# ==========================================
env_key = os.getenv("GEMINI_API_KEY")

secrets_key = None
try:
    if "GEMINI_API_KEY" in st.secrets:
        secrets_key = st.secrets["GEMINI_API_KEY"]
except Exception:
    pass

api_key = env_key or secrets_key

# Check if the API key is configured and not a placeholder
is_key_valid = False
if api_key and api_key.strip() != "" and "YOUR_GEMINI_API_KEY" not in api_key:
    is_key_valid = True

# ==========================================
# 5. MAIN AREA PANEL (REDESIGNED)
# ==========================================

# Centered glow blobs for visual aesthetics
st.markdown("""
<div class="glow-blob blob-1"></div>
<div class="glow-blob blob-2"></div>
""", unsafe_allow_html=True)

# Gradient header title (Centered)
st.markdown('<h1 class="glow-header" style="text-align: center; margin-top: 20px; margin-bottom: 5px; font-size: 46px;">🩺 Dr.Doc</h1>', unsafe_allow_html=True)
st.markdown('<p style="text-align: center; color: #94a3b8; font-size: 16px; margin-bottom: 30px;">Direct, secure PDF & DOCX intelligence platform. Zero setup required.</p>', unsafe_allow_html=True)

# 5a. API Key Error Block
if not is_key_valid:
    st.markdown("""
    <div class="welcome-card" style="border-color: rgba(239, 68, 68, 0.4); max-width: 650px;">
        <div style="font-size: 60px; margin-bottom: 15px;">🔑</div>
        <h2 style="color: #ef4444; font-size: 24px; font-weight: 600; margin-bottom: 10px;">API Key Required</h2>
        <p style="color: #94a3b8; font-size: 14px; max-width: 500px; margin: 0 auto 25px auto; line-height: 1.6;">
            Dr.Doc runs using automatic credentials. To use the application, please set your Gemini API key in the server configuration.
        </p>
        <div style="text-align: left; max-width: 450px; margin: 0 auto; background: rgba(0, 0, 0, 0.3); padding: 18px 24px; border-radius: 14px; border: 1px solid rgba(255, 255, 255, 0.05); font-family: monospace; font-size: 13px; color: #cbd5e1; line-height: 1.5;">
            <strong>Local Setup:</strong><br>
            Create a file named <code>.env</code> in the project directory with your key:<br>
            <code style="color: #c084fc; font-weight: bold;">GEMINI_API_KEY=AIzaSy...</code><br><br>
            <strong>Cloud Setup:</strong><br>
            Add the <code>GEMINI_API_KEY</code> variable to your Streamlit Cloud Secrets.
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# Render the file uploader at the top level on every run
uploaded_file = st.file_uploader(
    "Upload PDF or DOCX file",
    type=["pdf", "docx"],
    key="doc_uploader",
    label_visibility="collapsed" if st.session_state.current_file_name else "visible"
)

# 5b. Main layout logic based on uploaded file
if not uploaded_file:
    # Clear session values on file removal/emptying
    st.session_state.current_file_name = None
    st.session_state.document_text = ""
    st.session_state.document_summary = ""
    st.session_state.messages = []
    st.session_state.chat_session = None

    # Render welcome layout
    st.markdown("""
    <div class="welcome-card" style="margin-top: 15px;">
        <div style="font-size: 50px; margin-bottom: 10px;">📥</div>
        <h3 style="color: #f1f5f9; font-size: 22px; font-weight: 600; margin-bottom: 20px;">Analyze Your Document</h3>
        <p style="color: #94a3b8; font-size: 14px; max-width: 500px; margin: 0 auto 30px auto; line-height: 1.6;">
            Dr.Doc parses your document locally in the browser/server and uses Gemini's next-gen 1-million token context window to perform comprehensive analysis and exact-context chat.
        </p>
        <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; text-align: left; max-width: 700px; margin: 0 auto; border-top: 1px solid rgba(255, 255, 255, 0.05); padding-top: 25px;">
            <div style="background: rgba(255,255,255,0.02); padding: 15px; border-radius: 12px; border: 1px solid rgba(255,255,255,0.03);">
                <h5 style="color: #c084fc; margin-top:0; margin-bottom: 6px;">⚡ Direct Summary</h5>
                <p style="color: #64748b; font-size: 12px; margin: 0; line-height: 1.4;">Extract themes, highlights, and action lists automatically within seconds.</p>
            </div>
            <div style="background: rgba(255,255,255,0.02); padding: 15px; border-radius: 12px; border: 1px solid rgba(255,255,255,0.03);">
                <h5 style="color: #c084fc; margin-top:0; margin-bottom: 6px;">💬 Natural Chat</h5>
                <p style="color: #64748b; font-size: 12px; margin: 0; line-height: 1.4;">Stateful chatbot understands the entire document context accurately.</p>
            </div>
            <div style="background: rgba(255,255,255,0.02); padding: 15px; border-radius: 12px; border: 1px solid rgba(255,255,255,0.03);">
                <h5 style="color: #c084fc; margin-top:0; margin-bottom: 6px;">🔒 Pure Privacy</h5>
                <p style="color: #64748b; font-size: 12px; margin: 0; line-height: 1.4;">No third-party middleware. All file processing runs directly on your instance.</p>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

else:
    # File is uploaded. Reset state if the file itself changed.
    if st.session_state.current_file_name != uploaded_file.name:
        st.session_state.current_file_name = uploaded_file.name
        st.session_state.document_text = ""
        st.session_state.document_summary = ""
        st.session_state.messages = []
        st.session_state.chat_session = None

    # Render top bar with file stats and clean reset button
    st.markdown('<div class="glass-card" style="padding: 20px;">', unsafe_allow_html=True)
    col_file, col_stats, col_reset = st.columns([5, 5, 2])
    
    with col_file:
        st.markdown('<h4 style="font-size: 13px; margin-top: 0; color: #94a3b8; margin-bottom: 6px;">📂 Loaded Document</h4>', unsafe_allow_html=True)
        st.markdown(f"""
        <div style="background: rgba(138, 43, 226, 0.06); border: 1px solid rgba(138, 43, 226, 0.15); padding: 10px 14px; border-radius: 10px; font-size: 14px; font-weight: 600; color: #ffffff;">
            📄 {uploaded_file.name}
        </div>
        """, unsafe_allow_html=True)
        
    with col_stats:
        st.markdown('<h4 style="font-size: 13px; margin-top: 0; color: #94a3b8; margin-bottom: 6px;">📊 Quick Stats</h4>', unsafe_allow_html=True)
        if st.session_state.document_text:
            text = st.session_state.document_text
            words = len(text.split())
            read_time = math.ceil(words / 200)
            file_size_kb = len(uploaded_file.getvalue()) / 1024
            
            st.markdown(f"""
            <div style="display: flex; gap: 10px;">
                <div style="flex:1; background: rgba(255,255,255,0.02); border: 1px solid rgba(255,255,255,0.05); padding: 8px; border-radius: 8px; text-align:center;">
                    <span style="font-size:10px; color:#64748b; text-transform:uppercase;">Words</span><br>
                    <span style="font-size:15px; font-weight:700; color:#c084fc;">{words:,}</span>
                </div>
                <div style="flex:1; background: rgba(255,255,255,0.02); border: 1px solid rgba(255,255,255,0.05); padding: 8px; border-radius: 8px; text-align:center;">
                    <span style="font-size:10px; color:#64748b; text-transform:uppercase;">Read Time</span><br>
                    <span style="font-size:15px; font-weight:700; color:#c084fc;">{read_time} min</span>
                </div>
                <div style="flex:1; background: rgba(255,255,255,0.02); border: 1px solid rgba(255,255,255,0.05); padding: 8px; border-radius: 8px; text-align:center;">
                    <span style="font-size:10px; color:#64748b; text-transform:uppercase;">Size</span><br>
                    <span style="font-size:15px; font-weight:700; color:#c084fc;">{file_size_kb:.1f} KB</span>
                </div>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown("<div style='color:#64748b; font-size:13px; padding-top:8px;'>Extracting metrics...</div>", unsafe_allow_html=True)
            
    with col_reset:
        st.markdown('<h4 style="font-size: 13px; margin-top: 0; color: #94a3b8; margin-bottom: 6px;">⚙️ Controls</h4>', unsafe_allow_html=True)
        if st.button("🗑️ Reset", use_container_width=True):
            st.session_state.current_file_name = None
            st.session_state.document_text = ""
            st.session_state.document_summary = ""
            st.session_state.messages = []
            st.session_state.chat_session = None
            if "doc_uploader" in st.session_state:
                del st.session_state["doc_uploader"]
            if "gemini_client" in st.session_state:
                del st.session_state["gemini_client"]
            if "current_api_key" in st.session_state:
                del st.session_state["current_api_key"]
            st.rerun()
            
    st.markdown('</div>', unsafe_allow_html=True)

    # Initialize client in session state to persist it across runs and prevent connection closing
    if "gemini_client" not in st.session_state or st.session_state.get("current_api_key") != api_key:
        st.session_state.gemini_client = get_gemini_client(api_key)
        st.session_state.current_api_key = api_key
    client = st.session_state.gemini_client
    
    # 5c. Text Extraction Phase
    if not st.session_state.document_text:
        with st.status("🔮 Analyzing layout and extracting contents...", expanded=True) as status:
            if uploaded_file.name.lower().endswith(".pdf"):
                text = extract_text_from_pdf(uploaded_file)
            else:
                text = extract_text_from_docx(uploaded_file)
            
            if text:
                st.session_state.document_text = text
                status.update(label="✅ Document layout extracted successfully!", state="complete", expanded=False)
                st.rerun()
            else:
                status.update(label="❌ Failed to parse document content.", state="error")
                st.stop()
                
    # 5d. Summary Generation Phase
    if not st.session_state.document_summary:
        with st.spinner("🧠 Synthesizing executive summary & mapping concepts..."):
            summary = generate_summary(st.session_state.document_text, client)
            if summary:
                st.session_state.document_summary = summary
                st.rerun()
            else:
                st.error("Could not generate summary. Check your API key or network connection.")
                st.stop()

    # 5e. Chat Initialization Phase
    if not st.session_state.chat_session:
        try:
            # Start a chat session with the document pre-loaded
            st.session_state.chat_session = client.chats.create(
                model="gemini-3.5-flash",
                config=types.GenerateContentConfig(
                    system_instruction=(
                        "You are Dr.Doc, a premium AI document assistant. "
                        "You answer user questions strictly and accurately based on the content of the uploaded document. "
                        "If the answer cannot be found in the document, reply: 'I cannot find that information in the uploaded document. "
                        "Please ask a question related to its content.' Use clear formatting and Markdown in your responses."
                    )
                ),
                history=[
                    types.Content(
                        role="user", 
                        parts=[types.Part.from_text(text=f"Here is the content of the uploaded document:\n\n{st.session_state.document_text}")]
                    ),
                    types.Content(
                        role="model", 
                        parts=[types.Part.from_text(text="I have successfully parsed and analyzed the document. I am ready to answer any questions you have based on its content. What would you like to know?")]
                    )
                ]
            )
            st.session_state.messages = [
                {"role": "assistant", "content": "I have successfully parsed and analyzed the document. I am ready to answer any questions you have based on its content. What would you like to know?"}
            ]
        except Exception as e:
            st.error(f"Error starting chat session: {e}")
            st.stop()

    # ==========================================
    # 6. MAIN WORKSPACE TABS (GLASS TAB PANELS)
    # ==========================================
    tab_summary, tab_chat, tab_raw = st.tabs([
        "📊 Executive Summary", 
        "💬 Interactive Chatbot", 
        "📄 Document Raw Text"
    ])
    
    # --- TAB 1: EXECUTIVE SUMMARY ---
    with tab_summary:
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        st.markdown(st.session_state.document_summary)
        
        # Download summary button
        st.download_button(
            label="📥 Download Executive Summary (Markdown)",
            data=st.session_state.document_summary,
            file_name=f"summary_{os.path.splitext(uploaded_file.name)[0]}.md",
            mime="text/markdown",
            use_container_width=True
        )
        st.markdown('</div>', unsafe_allow_html=True)
        
    # --- TAB 2: INTERACTIVE CHATBOT ---
    with tab_chat:
        st.markdown('<div class="glass-card" style="padding-bottom: 20px;">', unsafe_allow_html=True)
        
        # Preset Quick Queries
        st.markdown('<p style="color: #94a3b8; font-size: 13px; font-weight: 600; margin-bottom: 8px; text-transform: uppercase; letter-spacing: 0.5px;">💡 Suggested Queries</p>', unsafe_allow_html=True)
        
        quick_prompts = [
            "Summarize the document in 3 sentences.",
            "What are the major risks or issues mentioned?",
            "Identify the key action items and deadlines.",
            "Who are the key stakeholders involved?"
        ]
        
        # Layout buttons horizontally
        cols = st.columns(len(quick_prompts))
        selected_quick_prompt = None
        for i, prompt_text in enumerate(quick_prompts):
            if cols[i].button(prompt_text, key=f"quick_{i}", use_container_width=True):
                selected_quick_prompt = prompt_text
                
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Render chat messages inside a glass layout
        for msg in st.session_state.messages:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])
                
        # Handle chat query (either from quick prompt or user text input)
        user_query = st.chat_input("Ask a question about this document...")
        
        # Override query if quick prompt was clicked
        if selected_quick_prompt:
            user_query = selected_quick_prompt
            
        if user_query:
            # Display user message
            with st.chat_message("user"):
                st.markdown(user_query)
            st.session_state.messages.append({"role": "user", "content": user_query})
            
            # Streaming response container
            with st.chat_message("assistant"):
                message_placeholder = st.empty()
                full_response = ""
                
                try:
                    # Stream the response from Gemini
                    response_stream = st.session_state.chat_session.send_message_stream(user_query)
                    for chunk in response_stream:
                        full_response += chunk.text
                        message_placeholder.markdown(full_response + "▌")
                    message_placeholder.markdown(full_response)
                    st.session_state.messages.append({"role": "assistant", "content": full_response})
                    
                    # Refresh page to maintain scroll and button states if needed
                    if selected_quick_prompt:
                        st.rerun()
                except Exception as e:
                    st.error(f"Error calling Gemini: {e}")
                    
    # --- TAB 3: DOCUMENT RAW TEXT ---
    with tab_raw:
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        st.markdown('<h3 style="font-size: 18px; margin-top: 0; color: #e2e8f0; margin-bottom: 5px;">📄 Extracted Full Text</h3>', unsafe_allow_html=True)
        st.markdown('<p style="color: #94a3b8; font-size: 13px; margin-bottom: 15px;">Below is the complete text extracted from the document that is provided to the chatbot model.</p>', unsafe_allow_html=True)
        
        # Raw text scroll container
        st.markdown(f'<div class="raw-text-container">{st.session_state.document_text}</div>', unsafe_allow_html=True)
        
        # Download raw text button
        st.download_button(
            label="📥 Download Extracted Text (TXT)",
            data=st.session_state.document_text,
            file_name=f"text_{os.path.splitext(uploaded_file.name)[0]}.txt",
            mime="text/plain",
            use_container_width=True
        )
        st.markdown('</div>', unsafe_allow_html=True)
