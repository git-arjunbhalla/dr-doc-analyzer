import streamlit as st
import os
import docx
from pypdf import PdfReader
from google import genai
from google.genai import types
from dotenv import load_dotenv
import math

# Load environment variables for local development
load_dotenv()

# ==========================================
# 1. PAGE SETUP & CONFIGURATION
# ==========================================
st.set_page_config(
    page_title="AuraDoc AI - Document Analyzer & Chatbot",
    page_icon="🌌",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Premium Custom CSS Injection for Glassmorphism & Sleek Dark Mode
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
    
    /* Main Background & Accent glow */
    [data-testid="stAppViewContainer"] {
        background-color: #0b0d19;
        background-image: 
            radial-gradient(at 10% 10%, rgba(138, 43, 226, 0.15) 0px, transparent 50%),
            radial-gradient(at 90% 85%, rgba(0, 191, 255, 0.12) 0px, transparent 50%);
        background-attachment: fixed;
    }
    
    /* Sidebar styling */
    [data-testid="stSidebar"] {
        background-color: #06070d !important;
        border-right: 1px solid rgba(138, 43, 226, 0.2);
    }
    
    /* Glassmorphism elements */
    .glass-card {
        background: rgba(23, 26, 47, 0.55);
        backdrop-filter: blur(16px) saturate(180%);
        -webkit-backdrop-filter: blur(16px) saturate(180%);
        border-radius: 16px;
        border: 1px solid rgba(138, 43, 226, 0.2);
        padding: 24px;
        margin-bottom: 20px;
        box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.3);
    }
    
    /* Glowing headers */
    .glow-header {
        background: linear-gradient(135deg, #a78bfa 0%, #8b5cf6 50%, #ec4899 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 800;
        letter-spacing: -0.5px;
    }
    
    /* Welcome landing page illustration card */
    .welcome-card {
        border: 1px dashed rgba(138, 43, 226, 0.4);
        background: rgba(23, 26, 47, 0.3);
        border-radius: 20px;
        padding: 40px;
        text-align: center;
        margin: 40px 0;
    }
    
    /* Document statistics badge container */
    .stats-container {
        display: flex;
        flex-wrap: wrap;
        gap: 12px;
        margin-bottom: 25px;
    }
    
    .stats-badge {
        flex: 1;
        min-width: 120px;
        padding: 12px 16px;
        background: rgba(138, 43, 226, 0.08);
        border: 1px solid rgba(138, 43, 226, 0.25);
        border-radius: 12px;
        text-align: center;
        transition: all 0.3s ease;
    }
    
    .stats-badge:hover {
        background: rgba(138, 43, 226, 0.15);
        border-color: rgba(138, 43, 226, 0.5);
        transform: translateY(-2px);
    }
    
    .stats-label {
        font-size: 11px;
        color: #94a3b8;
        text-transform: uppercase;
        font-weight: 600;
        letter-spacing: 0.5px;
    }
    
    .stats-value {
        font-size: 20px;
        font-weight: 700;
        color: #f1f5f9;
        margin-top: 4px;
        font-family: 'Outfit', sans-serif;
    }
    
    /* Custom tabs styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: transparent;
    }
    
    .stTabs [data-baseweb="tab"] {
        background-color: rgba(23, 26, 47, 0.4);
        border: 1px solid rgba(138, 43, 226, 0.1);
        border-radius: 8px 8px 0 0;
        padding: 8px 16px;
        color: #94a3b8;
        font-weight: 600;
        transition: all 0.3s ease;
    }
    
    .stTabs [data-baseweb="tab"]:hover {
        color: #c084fc;
        background-color: rgba(138, 43, 226, 0.1);
    }
    
    .stTabs [aria-selected="true"] {
        background-color: rgba(138, 43, 226, 0.15) !important;
        border-color: rgba(138, 43, 226, 0.4) !important;
        color: #f1f5f9 !important;
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
        padding: 16px;
        background: #07080e;
        border-radius: 12px;
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
        for i, page in enumerate(reader.pages):
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
            model="gemini-2.5-flash",
            contents=prompt
        )
        return response.text
    except Exception as e:
        st.error(f"Failed to generate summary: {e}")
        return None

# ==========================================
# 4. SIDEBAR PANEL (CONFIG & UPLOAD)
# ==========================================
with st.sidebar:
    st.markdown('<h2 class="glow-header" style="font-size: 24px; margin-bottom: 5px;">🌌 AuraDoc AI</h2>', unsafe_allow_html=True)
    st.markdown('<p style="color: #94a3b8; font-size: 13px; margin-bottom: 25px;">Document Intelligence Platform</p>', unsafe_allow_html=True)
    
    st.markdown('<h3 style="font-size: 16px; margin-bottom: 10px; color: #e2e8f0;">🔑 Authentication</h3>', unsafe_allow_html=True)
    
    # Check for pre-configured key
    env_key = os.getenv("GEMINI_API_KEY")
    
    secrets_key = None
    try:
        if "GEMINI_API_KEY" in st.secrets:
            secrets_key = st.secrets["GEMINI_API_KEY"]
    except Exception:
        pass
        
    default_key = env_key or secrets_key
    
    if default_key:
        api_key_input = st.text_input(
            "Gemini API Key",
            value=default_key,
            type="password",
            help="Pre-loaded from environment / secrets. You can overwrite it here if needed."
        )
        st.success("API Key detected automatically.")
    else:
        api_key_input = st.text_input(
            "Gemini API Key",
            placeholder="AIzaSy...",
            type="password",
            help="Enter your Google Gemini API key. You can get one from Google AI Studio."
        )
        st.info("💡 You can get a free API key at [Google AI Studio](https://aistudio.google.com/)")

    st.markdown("---")
    
    st.markdown('<h3 style="font-size: 16px; margin-bottom: 10px; color: #e2e8f0;">📂 Document Upload</h3>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader(
        "Upload PDF or DOCX",
        type=["pdf", "docx"],
        help="Max file size ~50MB. Larger documents will fit easily into Gemini's 1-million token context window."
    )
    
    # Process file change or clear button
    if uploaded_file is not None:
        if st.session_state.current_file_name != uploaded_file.name:
            # File has changed, reset session variables
            st.session_state.current_file_name = uploaded_file.name
            st.session_state.document_text = ""
            st.session_state.document_summary = ""
            st.session_state.messages = []
            st.session_state.chat_session = None
            
        st.markdown("---")
        st.markdown('<h3 style="font-size: 16px; margin-bottom: 10px; color: #e2e8f0;">📊 Document Statistics</h3>', unsafe_allow_html=True)
        
        # Calculate stats dynamically if text is extracted
        if st.session_state.document_text:
            text = st.session_state.document_text
            words = len(text.split())
            chars = len(text)
            read_time = math.ceil(words / 200) # Assuming average 200 WPM
            file_size_kb = len(uploaded_file.getvalue()) / 1024
            
            st.markdown(f"""
            <div class="stats-container">
                <div class="stats-badge">
                    <div class="stats-label">Words</div>
                    <div class="stats-value">{words:,}</div>
                </div>
                <div class="stats-badge">
                    <div class="stats-label">Characters</div>
                    <div class="stats-value">{chars:,}</div>
                </div>
                <div class="stats-badge">
                    <div class="stats-label">Est. Read Time</div>
                    <div class="stats-value">{read_time}m</div>
                </div>
                <div class="stats-badge">
                    <div class="stats-label">File Size</div>
                    <div class="stats-value">{file_size_kb:.1f} KB</div>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
        if st.button("🗑️ Clear Document & Reset", use_container_width=True):
            st.session_state.current_file_name = None
            st.session_state.document_text = ""
            st.session_state.document_summary = ""
            st.session_state.messages = []
            st.session_state.chat_session = None
            st.rerun()

# ==========================================
# 5. MAIN AREA PANEL (RESULTS & VISUALS)
# ==========================================

# Gradient header title
st.markdown('<h1 class="glow-header" style="text-align: center; margin-top: 10px; margin-bottom: 5px; font-size: 42px;">🌌 AuraDoc AI</h1>', unsafe_allow_html=True)
st.markdown('<p style="text-align: center; color: #94a3b8; font-size: 16px; margin-bottom: 40px;">Upload a PDF or DOCX file to extract deep summaries and engage in intelligent, contextual chat.</p>', unsafe_allow_html=True)

if not uploaded_file:
    # Landing / Welcome Page when no document is uploaded
    st.markdown("""
    <div class="welcome-card">
        <div style="font-size: 60px; margin-bottom: 15px;">📥</div>
        <h2 style="color: #f1f5f9; font-size: 24px; font-weight: 600; margin-bottom: 10px;">Awaiting Document Upload</h2>
        <p style="color: #94a3b8; font-size: 15px; max-width: 550px; margin: 0 auto 30px auto; line-height: 1.6;">
            AuraDoc AI parses your document locally in the browser/server and uses Gemini's next-gen 1-million token context window to perform comprehensive analysis and exact-context chat.
        </p>
        <div style="display: flex; justify-content: center; gap: 40px; text-align: left; max-width: 600px; margin: 0 auto;">
            <div>
                <h4 style="color: #c084fc; margin-bottom: 5px;">1. Set your Key</h4>
                <p style="color: #64748b; font-size: 13px;">Provide your Gemini API key in the sidebar for secure processing.</p>
            </div>
            <div>
                <h4 style="color: #c084fc; margin-bottom: 5px;">2. Upload File</h4>
                <p style="color: #64748b; font-size: 13px;">Drop any PDF or DOCX document up to 50MB in size.</p>
            </div>
            <div>
                <h4 style="color: #c084fc; margin-bottom: 5px;">3. Chat & Analyze</h4>
                <p style="color: #64748b; font-size: 13px;">Read the structured executive summary and start asking questions.</p>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

elif not api_key_input:
    # File uploaded but API key missing
    st.warning("⚠️ Please provide a valid Gemini API Key in the sidebar to begin analyzing the document.")

else:
    # Initialize client and run analysis
    client = get_gemini_client(api_key_input)
    
    # 5a. Text Extraction Phase
    if not st.session_state.document_text:
        with st.status("🔮 Processing file & extracting text...", expanded=True) as status:
            if uploaded_file.name.lower().endswith(".pdf"):
                text = extract_text_from_pdf(uploaded_file)
            else:
                text = extract_text_from_docx(uploaded_file)
            
            if text:
                st.session_state.document_text = text
                status.update(label="✅ Document parsed successfully!", state="complete", expanded=False)
                st.rerun()
            else:
                status.update(label="❌ Failed to parse document content.", state="error")
                st.stop()
                
    # 5b. Summary Generation Phase
    if not st.session_state.document_summary:
        with st.spinner("🧠 Generating Executive Summary with Gemini AI..."):
            summary = generate_summary(st.session_state.document_text, client)
            if summary:
                st.session_state.document_summary = summary
                st.rerun()
            else:
                st.error("Could not generate summary. Check your API key or network connection.")
                st.stop()

    # 5c. Chat Initialization Phase
    if not st.session_state.chat_session:
        try:
            # Start a chat session with the document pre-loaded
            st.session_state.chat_session = client.chats.create(
                model="gemini-2.5-flash",
                config=types.GenerateContentConfig(
                    system_instruction=(
                        "You are AuraDoc, a premium AI document assistant. "
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
    # 6. MAIN WORKSPACE TABS
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
            label="📥 Download Summary (Markdown)",
            data=st.session_state.document_summary,
            file_name=f"summary_{os.path.splitext(uploaded_file.name)[0]}.md",
            mime="text/markdown",
            use_container_width=True
        )
        st.markdown('</div>', unsafe_allow_html=True)
        
    # --- TAB 2: INTERACTIVE CHATBOT ---
    with tab_chat:
        st.markdown('<div class="glass-card" style="padding-bottom: 10px;">', unsafe_allow_html=True)
        st.markdown('<h3 style="font-size: 18px; margin-top: 0; color: #e2e8f0; margin-bottom: 5px;">💬 Conversational Document Assistant</h3>', unsafe_allow_html=True)
        st.markdown('<p style="color: #94a3b8; font-size: 13px; margin-bottom: 15px;">Ask specific questions, locate clauses, or search for data in the document.</p>', unsafe_allow_html=True)
        
        # Render Suggestion Buttons
        st.markdown('<p style="color: #cbd5e1; font-size: 12px; font-weight: 600; margin-bottom: 4px;">Quick Queries:</p>', unsafe_allow_html=True)
        
        # Create standard queries that set session state query value
        quick_prompts = [
            "Summarize the document in 3 sentences.",
            "What are the major risks or issues mentioned?",
            "Identify the key action items and deadlines.",
            "Who are the key people, organizations or stakeholders?"
        ]
        
        # We can use columns to lay out buttons
        cols = st.columns(len(quick_prompts))
        selected_quick_prompt = None
        for i, prompt_text in enumerate(quick_prompts):
            if cols[i].button(prompt_text, key=f"quick_{i}", use_container_width=True):
                selected_quick_prompt = prompt_text
                
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Render chat messages
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
                    with st.session_state.chat_session.send_message_stream(user_query) as stream:
                        for chunk in stream:
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
