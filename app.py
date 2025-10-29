import streamlit as st
from sentence_transformers import SentenceTransformer
import base64

# --- MODIFIED: Use aliases to prevent import collisions ---
from template_2 import (
    extract_text_from_pdf as t2_extract_pdf,
    extract_text_from_docx as t2_extract_docx,
    gemini_prompt as t2_gemini_prompt,
    call_gemini_api as t2_call_gemini,
    convert_to_docx as t2_convert_to_docx,
)
from template_1 import (
    extract_text_from_pdf as t1_extract_pdf,
    extract_text_from_docx as t1_extract_docx,
    gemini_prompt as t1_gemini_prompt,
    call_gemini_api as t1_call_gemini,
    convert_to_docx as t1_convert_to_docx,
)

st.set_page_config(page_title="TalentTune", layout="wide")

def template_1():
    st.markdown("<h3 style='color: rgb(186, 43, 43);'> Format Resume to Company Template (Template 1)</h3>", unsafe_allow_html=True)
    uploaded_file = st.file_uploader("Upload Resume (PDF or DOCX)", type=["pdf", "docx"], key="formatter-1")

    if uploaded_file:
        if uploaded_file.type == "application/pdf":
            # Use t1 alias
            resume_text = t1_extract_pdf(uploaded_file)
        elif uploaded_file.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
            # Use t1 alias
            resume_text = t1_extract_docx(uploaded_file)
        else:
            st.error("Unsupported file type.")
            return

        st.subheader("Extracted Resume Text")
        st.write(resume_text[:1000] + "..." if len(resume_text) > 1000 else resume_text)

        # --- MODIFIED: Use unique session state key ---
        if "formatted_resume_1" not in st.session_state:
            st.session_state.formatted_resume_1 = ""

        # --- MODIFIED: Add unique key to button ---
        if st.button("Format Resume", key="format_btn_1"):
            with st.spinner("Formatting..."):
                # Use t1 aliases
                prompt = t1_gemini_prompt(resume_text)
                formatted_resume = t1_call_gemini(prompt)
                # Use t1 session state
                st.session_state.formatted_resume_1 = formatted_resume

        if st.session_state.formatted_resume_1:
            st.subheader("Formatted Resume")
            st.markdown(st.session_state.formatted_resume_1)

            # Use t1 alias
            file_data = t1_convert_to_docx(st.session_state.formatted_resume_1)
            
            st.download_button(
                "Download DOCX",
                file_data,
                "formatted_resume_T1.docx", # Changed file name
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

def template_2():
    st.markdown("<h3 style='color: rgb(186, 43, 43);'> Format Resume to Company Template (Template 2)</h3>", unsafe_allow_html=True)
    uploaded_file = st.file_uploader("Upload Resume (PDF or DOCX)", type=["pdf", "docx"], key="formatter-2")

    if uploaded_file:
        if uploaded_file.type == "application/pdf":
            # Use t2 alias
            resume_text = t2_extract_pdf(uploaded_file)
        elif uploaded_file.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
            # Use t2 alias
            resume_text = t2_extract_docx(uploaded_file)
        else:
            st.error("Unsupported file type.")
            return

        st.subheader("Extracted Resume Text")
        st.write(resume_text[:1000] + "..." if len(resume_text) > 1000 else resume_text)

        # --- MODIFIED: Use unique session state key ---
        if "formatted_resume_2" not in st.session_state:
            st.session_state.formatted_resume_2 = ""

        # --- MODIFIED: Add unique key to button ---
        if st.button("Format Resume", key="format_btn_2"):
            with st.spinner("Formatting..."):
                # Use t2 aliases
                prompt = t2_gemini_prompt(resume_text)
                formatted_resume = t2_call_gemini(prompt)
                # Use t2 session state
                st.session_state.formatted_resume_2 = formatted_resume

        if st.session_state.formatted_resume_2:
            st.subheader("Formatted Resume")
            st.markdown(st.session_state.formatted_resume_2)

            # Use t2 alias
            file_data = t2_convert_to_docx(st.session_state.formatted_resume_2)
            
            st.download_button(
                "Download DOCX",
                file_data,
                "formatted_resume_T2.docx", # Changed file name
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

def set_bg_hack(main_bg):
    '''
    A function to unpack an image from root folder and set as bg.
 
    Returns
    -------
    The background.
    '''
    # set bg name
    main_bg_ext = "jpg"
        
    try:
        st.markdown(
             f"""
             <style>
             .stApp {{
                 background: url(data:image/{main_bg_ext};base64,{base64.b64encode(open(main_bg, "rb").read()).decode()});
                 background-size: cover
             }}
             </style>
             """,
             unsafe_allow_html=True
         )
    except FileNotFoundError:
        st.warning(f"Background image '{main_bg}' not found. Skipping background.")


def main():
    
    set_bg_hack('bg_final.jpg')  
    st.markdown(
    "<h1 style='color:rgb(186, 43, 43); font-weight: bold; text-align: center;'>TalentTune</h1>",
    unsafe_allow_html=True
)
    st.markdown(
    "<p style='text-align: center; font-size:18px;'><em>Fine-tune your profile for job success.</em></p>",
    unsafe_allow_html=True
)



    tab1, tab2 = st.tabs(["Template 1", "Template 2"]) # Renamed tabs for clarity

    # Footer
    st.markdown("<hr style='margin-top: 50px;'>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: grey;'>Powered by ModelMinds</p>", unsafe_allow_html=True)

    with tab1:
        template_1()

    with tab2:
        template_2()

if __name__ == "__main__":
    main()