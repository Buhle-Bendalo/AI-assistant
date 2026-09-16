import os
import shutil
import time
import uuid
import streamlit as st
from agent import SurcotecAgent
from logger import log_session
from tools import process_file_to_text

# ==========================================
# 1. CONFIGURATION & UI SETUP
# ==========================================
# Securely load API Key from Streamlit Secrets or Environment Variables
MY_API_KEY = st.secrets.get("GEMINI_API_KEY") or os.environ.get(
    "GEMINI_API_KEY", ""
)

OUTPUT_DIR = "generated_docs"
EXAMPLES_DIR = "Examples"

if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)

st.set_page_config(
    page_title="Surcotec Architect", page_icon="🏗️", layout="wide"
)

# Custom CSS for Professional Industrial Branding
st.markdown(
    """

    """,
    unsafe_allow_html=True,
)

# --- Initialize Session States ---
if "agent" not in st.session_state:
    st.session_state.agent = SurcotecAgent(MY_API_KEY)

if "messages" not in st.session_state:
    st.session_state.messages = []

if "session_id" not in st.session_state:
    st.session_state.session_id = str(uuid.uuid4())

if "uploaded_filename" not in st.session_state:
    st.session_state.uploaded_filename = "N/A"

if "user_corrections" not in st.session_state:
    st.session_state.user_corrections = []

# ==========================================
# 2. SIDEBAR - CONTROL PANEL
# ==========================================
with st.sidebar:
    st.title("🏗️ Surcotec Ops")
    st.caption("Intelligent Document Control")
    st.caption(
        "⚠️ Outputs are AI-generated and require human verification before use."
    )
    st.divider()

    if os.path.exists(EXAMPLES_DIR) and len(os.listdir(EXAMPLES_DIR)) > 0:
        st.write("✅ **Learning Examples Active**")
    else:
        st.warning("⚠️ No Examples folder found")

    st.subheader("📁 1. Upload Document")
    uploaded_file = st.file_uploader("Drop quote")

    if uploaded_file:
        st.session_state.uploaded_filename = uploaded_file.name
        st.success(f"File Ready: {uploaded_file.name}")

        if st.button("⚙️ Generate"):
            with st.spinner("Processing document and preparing job pack..."):
                file_bytes = uploaded_file.getvalue()
                extracted_text = process_file_to_text(
                    file_bytes, uploaded_file.name
                )

                prompt = (
                    f"I have uploaded a new document: {uploaded_file.name}. "
                    f"Content: \n{extracted_text}\n\n"
                    "Step 1: Compare this to the Master Template and Examples. "
                    "Step 2: List the proposed changes clearly for me to review. "
                    "Step 3: Call generate_excel_workbook to create the final job pack."
                )
                response = st.session_state.agent.ask(prompt)
                st.session_state.messages.append(
                    {"role": "assistant", "content": response}
                )

                log_session(
                    session_id=st.session_state.session_id,
                    uploaded_filename=uploaded_file.name,
                    extracted_fields={"raw_prompt_length": len(prompt)},
                    user_corrections=st.session_state.user_corrections,
                    output_filename="Pending output / Checked via agent response",
                )
                st.rerun()

    st.divider()
    st.subheader("📥 2. Download Results")

    if os.path.exists(OUTPUT_DIR):
        files = [f for f in os.listdir(OUTPUT_DIR) if f.endswith(".xlsx")]
        files.sort(
            key=lambda x: os.path.getmtime(os.path.join(OUTPUT_DIR, x)),
            reverse=True,
        )

        if files:
            latest = files[0]
            latest_path = os.path.join(OUTPUT_DIR, latest)
            with open(latest_path, "rb") as f_ptr:
                st.download_button(
                    label="📊 Download Latest",
                    data=f_ptr.read(),
                    file_name=latest,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_latest_{os.path.getmtime(latest_path)}",
                    use_container_width=True,
                )

            if len(files) > 1:
                with st.expander("🗂️ Download older files"):
                    selected = st.selectbox(
                        "Select a file",
                        options=files[1:],
                        label_visibility="collapsed",
                    )
                    selected_path = os.path.join(OUTPUT_DIR, selected)
                    with open(selected_path, "rb") as f_ptr:
                        st.download_button(
                            label=f"⬇️ Download: {selected[:22]}",
                            data=f_ptr.read(),
                            file_name=selected,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"dl_old_{selected}_{os.path.getmtime(selected_path)}",
                            use_container_width=True,
                        )

            if st.button("🚨 Clear File History", use_container_width=True):
                for f in files:
                    os.remove(os.path.join(OUTPUT_DIR, f))
                st.toast("Cleared!", icon="🔥")
                time.sleep(1)
                st.rerun()
        else:
            st.info("No files ready.")

    st.divider()
    if st.button("🗑️ Clear Chat", use_container_width=True):
        st.session_state.messages = []
        st.session_state.user_corrections = []
        st.rerun()

# ==========================================
# 3. MAIN CHAT INTERFACE
# ==========================================
st.title("Surcotec Document Architect")
st.warning(
    "⚠️ **Disclaimer:** All outputs are AI-generated and must be verified by human personnel prior to release."
)
st.markdown("---")

col_chat, col_status = st.columns([3, 1])

with col_chat:
    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    if user_query := st.chat_input(
            "Type in corrections or type produce/proceed..."
    ):
        st.session_state.messages.append(
            {"role": "user", "content": user_query}
        )
        st.session_state.user_corrections.append(user_query)

        with st.chat_message("user"):
            st.markdown(user_query)

        with st.chat_message("assistant"):
            with st.spinner("Processing..."):
                response = st.session_state.agent.ask(user_query)
                st.markdown(response)
                st.session_state.messages.append(
                    {"role": "assistant", "content": response}
                )

                if "SUCCESS" in response:
                    log_session(
                        session_id=st.session_state.session_id,
                        uploaded_filename=st.session_state.uploaded_filename,
                        extracted_fields={"chat_status": "Success generated"},
                        user_corrections=st.session_state.user_corrections,
                        output_filename="Generated successfully",
                    )
                    st.toast("Excel Generated!", icon="📊")
                    time.sleep(1.5)
                    st.rerun()

with col_status:
    st.subheader("System Intel")
    st.info("🤖 **Engine:** Gemini")
    st.success("📝 **Template:** Active")

    with st.expander("Applying Corrections"):
        st.write("""
        If you see a mistake in the preview:
        1. Tell the bot: *"Change X to Y"*
        2. Wait for confirmation.
        3. Type **produce** or **proceed** to generate.
        """)