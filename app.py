import streamlit as st
import requests
from requests.adapters import HTTPAdapter
from urllib3.util import Retry

st.set_page_config(page_title="Outlook Bulk Reader", layout="wide")
st.title("📬 Outlook Bulk Mailbox Reader")

if "accounts" not in st.session_state:
    st.session_state.accounts = {}

# Configure a robust session with retries for handling server network drops
session = requests.Session()
retries = Retry(total=3, backoff_factor=1, status_forcelist=[500, 502, 503, 504])
session.mount('https://', HTTPAdapter(max_retries=retries))

# ------------------ SIDEBAR: INPUT CREDENTIALS ------------------
st.sidebar.header("🔑 Credentials Input")
raw_input = st.sidebar.text_area(
    "Enter lines (Format: mail:pass:refresh_token:client_id)",
    height=250,
    placeholder="email1@outlook.com:pass:token:client_id"
)

if st.sidebar.button("🚀 Load & Parse Accounts", use_container_width=True):
    lines = raw_input.strip().split("\n")
    parsed_accounts = {}
    for line in lines:
        if not line.strip():
            continue
        parts = line.strip().split(":")
        if len(parts) >= 4:
            email = parts[0]
            # Capture components properly even if values contain extra colons
            parsed_accounts[email] = {
                "refresh_token": parts[2],
                "client_id": parts[3]
            }
    st.session_state.accounts = parsed_accounts
    st.sidebar.success(f"Loaded {len(parsed_accounts)} accounts!")

# ------------------ MAIN SECTION: INBOX ------------------
if not st.session_state.accounts:
    st.info("💡 Please enter your credentials in the sidebar and click 'Load & Parse Accounts'.")
else:
    selected_email = st.selectbox("🎯 Select Account to Read:", list(st.session_state.accounts.keys()))
    
    if selected_email:
        st.subheader(f"📨 Inbox for {selected_email}")
        account_data = st.session_state.accounts[selected_email]
        
        # Explicitly targeted production login domain
        token_url = "https://microsoftonline.com"
        token_data = {
            "client_id": account_data["client_id"],
            "grant_type": "refresh_token",
            "refresh_token": account_data["refresh_token"],
            "scope": "https://microsoft.com"
        }
        
        with st.spinner("Exchanging token and fetching emails..."):
            try:
                token_response = session.post(token_url, data=token_data, timeout=15)
                token_json = token_response.json()
                
                if "access_token" not in token_json:
                    st.error(f"❌ Failed to get Access Token. Response:\n{token_json}")
                else:
                    access_token = token_json["access_token"]
                    
                    graph_url = f"https://microsoft.com{selected_email}/mailFolders/inbox/messages?$top=15"
                    headers = {
                        "Authorization": f"Bearer {access_token}",
                        "Content-Type": "application/json"
                    }
                    
                    mail_response = session.get(graph_url, headers=headers, timeout=15)
                    
                    if mail_response.status_code == 200:
                        emails_list = mail_response.json().get("value", [])
                        
                        if not emails_list:
                            st.warning("📭 This inbox is completely empty.")
                        else:
                            for idx, mail in enumerate(emails_list, 1):
                                subject = mail.get("subject", "No Subject")
                                sender = mail.get("from", {}).get("emailAddress", {}).get("address", "Unknown")
                                date = mail.get("receivedDateTime", "Unknown Date")
                                preview = mail.get("bodyPreview", "No preview available.")
                                
                                with st.expander(f"**[{idx}] From:** {sender} | **Subject:** {subject}"):
                                    st.write(f"📅 **Date:** {date}")
                                    st.write(f"📝 **Preview:** {preview}")
                    else:
                        st.error(f"❌ Microsoft API Error ({mail_response.status_code}): {mail_response.text}")
                        
            except Exception as e:
                st.error(f"💥 Network connection error occurred: {str(e)}")
