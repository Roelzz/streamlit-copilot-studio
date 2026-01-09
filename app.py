"""
Copilot Studio Chat Interface
A Streamlit chat app using the M365 Agents SDK.
"""

import asyncio
import os
import streamlit as st
from streamlit_msal import Msal
from dotenv import load_dotenv

from copilot_client import CopilotStudioClient, clean_citations, format_references_html, sanitize_html

load_dotenv()


def render_adaptive_card_element(element, depth=0):
    """Recursively render an adaptive card element."""
    if not isinstance(element, dict):
        return

    elem_type = element.get('type', '')

    if elem_type == 'TextBlock':
        text = element.get('text', '')
        weight = element.get('weight', 'default')
        size = element.get('size', 'default')
        horizontal_alignment = element.get('horizontalAlignment', 'left')
        is_subtle = element.get('isSubtle', False)

        # Apply formatting
        if not text:
            return

        # Size mapping
        size_map = {
            'Small': '0.9em',
            'Default': '1em',
            'Medium': '1.2em',
            'Large': '1.5em',
            'ExtraLarge': '2em'
        }
        font_size = size_map.get(size, '1em')

        # Build markdown with HTML styling
        style = f"font-size: {font_size};"
        if horizontal_alignment.lower() == 'center':
            style += " text-align: center;"
        elif horizontal_alignment.lower() == 'right':
            style += " text-align: right;"
        if is_subtle:
            style += " opacity: 0.7;"

        if weight == 'Bolder' or weight == 'bolder':
            st.markdown(f'<div style="{style}"><strong>{text}</strong></div>', unsafe_allow_html=True)
        else:
            st.markdown(f'<div style="{style}">{text}</div>', unsafe_allow_html=True)

    elif elem_type == 'Image':
        url = element.get('url', '')
        alt_text = element.get('altText', '')
        size = element.get('size', 'Auto')

        if url:
            # Size mapping for images
            if size == 'Small':
                st.image(url, width=80, caption=alt_text if alt_text else None)
            elif size == 'Medium':
                st.image(url, width=120, caption=alt_text if alt_text else None)
            elif size == 'Large':
                st.image(url, width=200, caption=alt_text if alt_text else None)
            else:  # Auto, Stretch
                st.image(url, caption=alt_text if alt_text else None)

    elif elem_type == 'Container':
        items = element.get('items', [])
        # Render container items in a styled div
        st.markdown('<div style="padding: 10px; margin: 5px 0;">', unsafe_allow_html=True)
        for item in items:
            render_adaptive_card_element(item, depth + 1)
        st.markdown('</div>', unsafe_allow_html=True)

    elif elem_type == 'ColumnSet':
        columns = element.get('columns', [])
        if columns:
            cols = st.columns(len(columns))
            for idx, column in enumerate(columns):
                with cols[idx]:
                    items = column.get('items', [])
                    for item in items:
                        render_adaptive_card_element(item, depth + 1)

    elif elem_type == 'ProgressBar':
        # Simple progress bar representation
        st.markdown('<div style="background: #e0e0e0; height: 4px; border-radius: 2px; margin: 10px 0;"><div style="background: #1976d2; height: 4px; width: 60%; border-radius: 2px;"></div></div>', unsafe_allow_html=True)

    elif elem_type == 'ActionSet':
        actions = element.get('actions', [])
        horizontal_alignment = element.get('horizontalAlignment', 'left')

        if actions:
            # Create button layout
            align_style = ""
            if horizontal_alignment.lower() == 'center':
                align_style = "text-align: center;"
            elif horizontal_alignment.lower() == 'right':
                align_style = "text-align: right;"

            st.markdown(f'<div style="margin: 10px 0; {align_style}">', unsafe_allow_html=True)
            button_html = ""
            for action in actions:
                title = action.get('title', '')
                action_type = action.get('type', '')
                if title:
                    button_html += f'<button style="margin: 0 5px; padding: 8px 16px; border: 1px solid #ccc; border-radius: 4px; background: #f5f5f5; cursor: pointer;">{title}</button>'
            st.markdown(button_html, unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

# Validate required environment variables at startup
REQUIRED_ENV_VARS = {
    "COPILOT_ENVIRONMENT_ID": "Copilot Studio environment ID",
    "COPILOT_AGENT_IDENTIFIER": "Copilot Studio agent identifier",
    "AZURE_TENANT_ID": "Azure tenant ID",
    "AZURE_APP_CLIENT_ID": "Azure app client ID",
}

missing_vars = []
for var_name, var_description in REQUIRED_ENV_VARS.items():
    if not os.getenv(var_name):
        missing_vars.append(f"- **{var_name}**: {var_description}")

if missing_vars:
    # Page config needs to be set first even for error pages
    st.set_page_config(
        page_title="Configuration Error",
        page_icon="⚠️",
        layout="centered",
    )
    st.error("Missing required environment variables")
    st.markdown("Please configure the following in your `.env` file:\n\n" + "\n".join(missing_vars))
    st.info("Copy `.env.example` to `.env` and fill in your values. See README.md for details.")
    st.stop()

# Page config
st.set_page_config(
    page_title="Copilot Studio",
    page_icon="💬",
    layout="centered",
)

# Minimal styling
st.markdown("""
<style>
    header {visibility: hidden;}
    .block-container {padding-top: 2rem;}
</style>
""", unsafe_allow_html=True)


def init_session():
    """Initialize session state."""
    if "messages" not in st.session_state:
        st.session_state.messages = []
    if "client" not in st.session_state:
        st.session_state.client = None


def main():
    init_session()

    st.title("💬 Copilot Studio")

    # Authentication via streamlit-msal
    auth_data = Msal.initialize_ui(
        client_id=os.getenv("AZURE_APP_CLIENT_ID"),
        authority=f"https://login.microsoftonline.com/{os.getenv('AZURE_TENANT_ID')}",
        scopes=["https://api.powerplatform.com/.default"],
        sign_in_label="Sign in with Microsoft",
        sign_out_label="Sign out",
    )

    if not auth_data:
        st.info("Please sign in to chat with Copilot Studio.")
        st.stop()

    # New Chat button in header area
    col1, col2 = st.columns([6, 1])
    with col2:
        if st.button("🗑️ New", help="Start a new conversation"):
            st.session_state.messages = []
            st.session_state.client = None
            st.rerun()

    # Get access token
    access_token = auth_data.get("accessToken")
    if not access_token:
        st.error("Failed to get access token.")
        st.stop()

    # Initialize client if needed
    if st.session_state.client is None:
        with st.spinner("Connecting to Copilot Studio..."):
            try:
                client = CopilotStudioClient(access_token)
                # Add timeout for initial connection (30 seconds)
                welcome = asyncio.run(
                    asyncio.wait_for(client.start_conversation(), timeout=30)
                )
                st.session_state.client = client

                if welcome:
                    st.session_state.messages.append({
                        "role": "assistant",
                        "content": welcome
                    })
            except asyncio.TimeoutError:
                st.error("Connection to Copilot Studio timed out.")
                st.info("Please check your network connection and try again.")
                st.stop()
            except Exception as e:
                st.error(f"Failed to connect to Copilot Studio: {str(e)}")
                st.info("Please check your configuration in `.env` and ensure the agent is published in Copilot Studio.")
                st.stop()

    # Display messages
    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            # Sanitize HTML before displaying (prevents XSS attacks)
            if msg["role"] == "assistant":
                sanitized_content = sanitize_html(msg["content"])
                st.markdown(sanitized_content, unsafe_allow_html=True)

                # Display adaptive cards if present
                if "adaptive_cards" in msg and msg["adaptive_cards"]:
                    import json
                    for idx, card in enumerate(msg["adaptive_cards"]):
                        with st.expander(f"📋 Adaptive Card {idx + 1}", expanded=False):
                            # Check if card is HTML string or JSON object
                            if isinstance(card, str):
                                # It's HTML content - sanitize and render
                                sanitized_card_html = sanitize_html(card)
                                st.markdown(sanitized_card_html, unsafe_allow_html=True)

                                # Also show raw HTML in a code block for debugging
                                with st.expander("View Raw HTML", expanded=False):
                                    st.code(card, language="html")
                            elif isinstance(card, dict):
                                # It's JSON - render as Adaptive Card
                                card_type = card.get('type', '')
                                if card_type == 'AdaptiveCard':
                                    # Render card with a styled container
                                    st.markdown('<div style="border: 1px solid #e0e0e0; border-radius: 8px; padding: 15px; background: white; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">', unsafe_allow_html=True)

                                    body = card.get('body', [])
                                    for element in body:
                                        render_adaptive_card_element(element)

                                    st.markdown('</div>', unsafe_allow_html=True)

                                # Show JSON structure in collapsible section
                                with st.expander("View JSON Structure", expanded=False):
                                    st.json(card)
                            else:
                                # Unknown format - just display it
                                st.write(card)
            else:
                # User messages should be plain text only
                st.markdown(msg["content"])

    # Chat input
    if prompt := st.chat_input("Message Copilot..."):
        # Add user message
        st.session_state.messages.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)

        # Stream response
        with st.chat_message("assistant"):
            # Container for reasoning/thoughts (collapsible)
            thinking_container = st.empty()
            status_placeholder = st.empty()
            content_placeholder = st.empty()

            async def process_response():
                content_parts = []
                suggestions = None
                citation_metadata = {}
                search_results = []  # Collect search results by index
                thoughts = []  # Collect chain-of-thought
                adaptive_cards = []  # Collect adaptive cards
                got_streaming = False

                try:
                    async for msg_type, msg_content in st.session_state.client.send_message(prompt):
                        if msg_type == 'status':
                            status_placeholder.caption(f"_{msg_content}_")
                        elif msg_type == 'thought':
                            # Collect reasoning/chain-of-thought
                            thoughts.append(msg_content)
                            # Update thinking display
                            with thinking_container.status("Thinking...", expanded=False) as status:
                                for t in thoughts:
                                    task = t.get('task', 'Processing')
                                    text = t.get('text', '')
                                    st.write(f"**{task}**: {text}")
                        elif msg_type == 'search_result':
                            # Collect search results (contain URLs)
                            search_results.append(msg_content)
                        elif msg_type == 'content':
                            got_streaming = True
                            content_parts.append(msg_content)
                            # Show accumulated content with citations cleaned (plain text during streaming)
                            accumulated = "".join(content_parts)
                            cleaned, _ = clean_citations(accumulated)
                            content_placeholder.markdown(cleaned)
                        elif msg_type == 'final_content':
                            # Non-streaming response - use this only if we didn't get streaming chunks
                            if not got_streaming:
                                content_parts = [msg_content]
                                cleaned, _ = clean_citations(msg_content)
                                content_placeholder.markdown(cleaned)
                        elif msg_type == 'citations':
                            # Merge citation metadata from entities
                            # Try to enrich with URLs from search results
                            for cite_id, cite_info in msg_content.items():
                                # Citation IDs are like 'turn52search0' - extract index
                                import re
                                match = re.search(r'search(\d+)$', cite_id)
                                if match and not cite_info.get('url'):
                                    idx = int(match.group(1))
                                    # Find matching search result by index
                                    for sr in search_results:
                                        if sr.get('index') == idx:
                                            cite_info['url'] = sr.get('url', '')
                                            if not cite_info.get('title'):
                                                cite_info['title'] = sr.get('title', '')
                                            break
                            citation_metadata.update(msg_content)
                        elif msg_type == 'adaptive_card':
                            # Collect adaptive cards
                            adaptive_cards.append(msg_content)
                        elif msg_type == 'attachment':
                            # Handle other attachments
                            pass  # Could expand this later
                        elif msg_type == 'suggestion':
                            suggestions = msg_content

                    # Finalize thinking display
                    if thoughts:
                        with thinking_container.status("Reasoning", expanded=False, state="complete") as status:
                            for t in thoughts:
                                task = t.get('task', 'Processing')
                                text = t.get('text', '')
                                st.write(f"**{task}**: {text}")

                    # Clear status when done
                    status_placeholder.empty()

                    # Return cleaned final content with clickable HTML citations
                    raw_content = "".join(content_parts)
                    cleaned_text, citations = clean_citations(raw_content, use_html=True, citation_metadata=citation_metadata)

                    # Add references section with clickable links
                    if citations:
                        cleaned_text += format_references_html(citations)

                    return cleaned_text, citations, suggestions, adaptive_cards

                except Exception as e:
                    status_placeholder.empty()
                    error_msg = f"Error during conversation: {str(e)}"
                    st.error(error_msg)
                    return f"Sorry, I encountered an error while processing your request. Please try again.", {}, None, []

            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            try:
                # Add timeout protection (5 minutes max)
                response, citations, suggestions, adaptive_cards = loop.run_until_complete(
                    asyncio.wait_for(process_response(), timeout=300)
                )
            except asyncio.TimeoutError:
                response = "Sorry, the request timed out. The agent took too long to respond."
                citations = {}
                suggestions = None
                adaptive_cards = []
                st.error("Please try again with a simpler question or start a new conversation.")
            except Exception as e:
                response = f"Sorry, an unexpected error occurred: {str(e)}"
                citations = {}
                suggestions = None
                adaptive_cards = []
                st.error("If this error persists, try starting a new conversation.")
            finally:
                loop.close()

            # Render final response (with clickable HTML citations if any)
            # Sanitize before displaying to prevent XSS
            sanitized_response = sanitize_html(response)
            content_placeholder.markdown(sanitized_response, unsafe_allow_html=True)

            # Render adaptive cards if any
            if adaptive_cards:
                import json
                for idx, card in enumerate(adaptive_cards):
                    with st.expander(f"📋 Adaptive Card {idx + 1}", expanded=True):
                        # Check if card is HTML string or JSON object
                        if isinstance(card, str):
                            # It's HTML content - sanitize and render
                            sanitized_card_html = sanitize_html(card)
                            st.markdown(sanitized_card_html, unsafe_allow_html=True)

                            # Also show raw HTML in a code block for debugging
                            with st.expander("View Raw HTML", expanded=False):
                                st.code(card, language="html")
                        elif isinstance(card, dict):
                            # It's JSON - render as Adaptive Card
                            card_type = card.get('type', '')
                            if card_type == 'AdaptiveCard':
                                # Render card with a styled container
                                st.markdown('<div style="border: 1px solid #e0e0e0; border-radius: 8px; padding: 15px; background: white; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">', unsafe_allow_html=True)

                                body = card.get('body', [])
                                for element in body:
                                    render_adaptive_card_element(element)

                                st.markdown('</div>', unsafe_allow_html=True)

                            # Show JSON structure in collapsible section
                            with st.expander("View JSON Structure", expanded=False):
                                st.json(card)
                        else:
                            # Unknown format - just display it
                            st.write(card)

            # Show suggestions if any
            if suggestions:
                st.caption(f"**Suggestions:** {suggestions}")

        # Store response with HTML citations and adaptive cards for history
        st.session_state.messages.append({
            "role": "assistant",
            "content": response,
            "adaptive_cards": adaptive_cards if adaptive_cards else []
        })


if __name__ == "__main__":
    main()
