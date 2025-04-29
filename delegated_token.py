"""
Streamlit-friendly helper:
▶ first run – shows a URL + device-code you open in a new tab
▶ later runs – silently refreshes using token_cache.json
"""

import os, msal, json, pathlib, streamlit as st

CLIENT_ID  = st.secrets.AZURE_AD.CLIENT_ID
TENANT_ID  = st.secrets.AZURE_AD.TENANT_ID
AUTHORITY  = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES     = [
    "Chat.Create",
    "ChatMessage.Send",
    "User.ReadBasic.All",
    "Directory.Read.All",
    "offline_access",
]
CACHE_FILE = pathlib.Path("token_cache.json")

def get_graph_token():
    cache = msal.SerializableTokenCache()
    if CACHE_FILE.exists():
        cache.deserialize(CACHE_FILE.read_text())

    app = msal.PublicClientApplication(
        CLIENT_ID, authority=AUTHORITY, token_cache=cache
    )

    # 1️⃣  Silent first
    accounts = app.get_accounts()
    result   = app.acquire_token_silent(SCOPES, account=accounts[0] if accounts else None)

    # 2️⃣  Device-code if no cached token
    if not result:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:                 # something went wrong
            st.error("Device flow error")
            st.stop()

        st.info(
            f"**One-time sign-in required**  \n"
            f"Open **{flow['verification_uri']}**  \n"
            f"Enter code **{flow['user_code']}**  \n"
            f"(sign in as the bot account, then come back)"
        )
        result = app.acquire_token_by_device_flow(flow)   # waits until auth completed

    # 3️⃣  Persist cache
    if cache.has_state_changed:
        CACHE_FILE.write_text(cache.serialize())

    return result["access_token"]
