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
        weight = element.get('weight') or 'default'
        size = element.get('size') or 'default'
        horizontal_alignment = element.get('horizontalAlignment') or 'left'
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

        # Build markdown with HTML styling - Editorial typography
        style = f"font-size: {font_size}; color: #292524; line-height: 1.65; margin-bottom: 0.75rem; font-family: 'Manrope', sans-serif;"
        if horizontal_alignment and horizontal_alignment.lower() == 'center':
            style += " text-align: center;"
        elif horizontal_alignment and horizontal_alignment.lower() == 'right':
            style += " text-align: right;"
        if is_subtle:
            style += " color: #6B7280;"

        if weight and (weight.lower() == 'bolder'):
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
        # Render container items in a styled div - Refined editorial
        st.markdown('<div style="padding: 1.25rem; margin: 1rem 0; background: linear-gradient(135deg, #FAFAF9 0%, #FFFFFF 100%); border-radius: 10px; border: 1.5px solid #E7E5E4; box-shadow: 0 2px 4px -1px rgba(0, 0, 0, 0.06);">', unsafe_allow_html=True)
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
        # Refined progress bar with animated gradient
        st.markdown('''<div style="
    background: linear-gradient(90deg, #E7E5E4 0%, #D6D3D1 100%);
    height: 10px;
    border-radius: 6px;
    margin: 1.25rem 0;
    overflow: hidden;
    box-shadow: inset 0 2px 4px rgba(0, 0, 0, 0.06);
"><div style="
    background: linear-gradient(90deg, #1D4ED8 0%, #3B82F6 50%, #10B981 100%);
    background-size: 200% 100%;
    animation: shimmer 2s linear infinite;
    height: 10px;
    width: 60%;
    border-radius: 6px;
    box-shadow: 0 2px 4px rgba(29, 78, 216, 0.4);
"></div></div>''', unsafe_allow_html=True)

    elif elem_type == 'FactSet':
        facts = element.get('facts', [])
        if facts:
            # Render facts as an editorial table with refined styling
            st.markdown('''<div style="
    margin: 1.25rem 0;
    border: 1.5px solid #E7E5E4;
    border-radius: 10px;
    overflow: hidden;
    background: linear-gradient(135deg, #FFFFFF 0%, #FAFAF9 100%);
    box-shadow: 0 2px 4px -1px rgba(0, 0, 0, 0.06);
">''', unsafe_allow_html=True)

            for idx, fact in enumerate(facts):
                title = fact.get('title', '')
                value = fact.get('value', '')
                bg_color = 'linear-gradient(90deg, #FAFAF9 0%, #FFFFFF 100%)' if idx % 2 == 0 else '#FFFFFF'
                border_style = '1px solid #E7E5E4' if idx < len(facts) - 1 else 'none'
                st.markdown(f'''<div style="
    display: flex;
    padding: 1rem 1.25rem;
    background: {bg_color};
    border-bottom: {border_style};
    transition: background 0.2s ease;
">
    <div style="
        flex: 0 0 40%;
        font-weight: 600;
        color: #57534E;
        font-size: 0.875rem;
        font-family: 'Manrope', sans-serif;
        letter-spacing: 0.01em;
    ">{title}</div>
    <div style="
        flex: 1;
        color: #292524;
        font-size: 0.875rem;
        font-family: 'Manrope', sans-serif;
    ">{value}</div>
</div>''', unsafe_allow_html=True)

            st.markdown('</div>', unsafe_allow_html=True)

    elif elem_type == 'ImageSet':
        images = element.get('images', [])
        image_size = element.get('imageSize', 'medium')
        if images:
            # Use Streamlit columns for image grid
            cols = st.columns(min(len(images), 4))  # Max 4 columns
            for idx, img in enumerate(images):
                with cols[idx % 4]:
                    url = img.get('url', '')
                    if url:
                        if image_size.lower() == 'small':
                            st.image(url, width=60)
                        elif image_size.lower() == 'large':
                            st.image(url, width=150)
                        else:  # medium
                            st.image(url, width=100)

    elif elem_type == 'RichTextBlock':
        inlines = element.get('inlines', [])
        if inlines:
            rich_text = ""
            for inline in inlines:
                if isinstance(inline, dict):
                    inline_type = inline.get('type', '')
                    if inline_type == 'TextRun':
                        text = inline.get('text', '')
                        color = inline.get('color')
                        weight = inline.get('weight')
                        size = inline.get('size')
                        italic = inline.get('italic', False)

                        # Build styled span
                        style = ""
                        if color:
                            style += f"color: {color};"
                        if weight == 'Bolder':
                            text = f"<strong>{text}</strong>"
                        if italic:
                            text = f"<em>{text}</em>"

                        rich_text += f'<span style="{style}">{text}</span>'
                elif isinstance(inline, str):
                    rich_text += inline

            if rich_text:
                st.markdown(f'''<div style="
    margin: 1rem 0;
    line-height: 1.7;
    color: #292524;
    font-size: 1rem;
    font-family: 'Manrope', sans-serif;
">{rich_text}</div>''', unsafe_allow_html=True)

    elif elem_type == 'Table':
        columns = element.get('columns', [])
        rows = element.get('rows', [])

        if columns and rows:
            # Build editorial HTML table with refined styling
            table_html = '''<table style="
    width: 100%;
    border-collapse: separate;
    border-spacing: 0;
    margin: 1.25rem 0;
    border-radius: 10px;
    overflow: hidden;
    box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.08);
    border: 1.5px solid #E7E5E4;
    font-family: 'Manrope', sans-serif;
">'''

            # Header row with sophisticated gradient
            table_html += '''<tr style="
    background: linear-gradient(135deg, #F5F5F4 0%, #E7E5E4 100%);
    border-bottom: 2px solid #D6D3D1;
">'''
            for col in columns:
                width = col.get('width', 'auto')
                table_html += '''<th style="
    padding: 1rem 1.25rem;
    text-align: left;
    font-weight: 700;
    font-size: 0.8125rem;
    color: #44403C;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    font-family: 'Manrope', sans-serif;
">'''
                table_html += '</th>'
            table_html += '</tr>'

            # Data rows with subtle alternating gradient backgrounds
            for row_idx, row in enumerate(rows):
                cells = row.get('cells', [])
                bg_color = 'linear-gradient(90deg, #FFFFFF 0%, #FAFAF9 100%)' if row_idx % 2 == 0 else '#FFFFFF'
                border_style = '1px solid #E7E5E4' if row_idx < len(rows) - 1 else 'none'
                table_html += f'<tr style="border-bottom: {border_style}; background: {bg_color}; transition: background 0.2s ease;">'
                for cell in cells:
                    table_html += '''<td style="
    padding: 0.875rem 1.25rem;
    color: #292524;
    font-size: 0.875rem;
    line-height: 1.6;
    font-family: 'Manrope', sans-serif;
">'''
                    # Recursively render cell items
                    items = cell.get('items', [])
                    for item in items:
                        # For now, just extract text
                        if isinstance(item, dict) and item.get('type') == 'TextBlock':
                            table_html += item.get('text', '')
                    table_html += '</td>'
                table_html += '</tr>'

            table_html += '</table>'
            st.markdown(table_html, unsafe_allow_html=True)

    elif elem_type.startswith('Input.'):
        # Render input elements (read-only in Streamlit)
        input_type = elem_type.replace('Input.', '')
        label = element.get('label', '') or element.get('placeholder', '')
        input_id = element.get('id', '')

        st.markdown(f'''<div style="margin: 1.25rem 0;">
    <label style="
        display: block;
        margin-bottom: 0.625rem;
        font-weight: 600;
        color: #44403C;
        font-size: 0.875rem;
        font-family: 'Manrope', sans-serif;
        letter-spacing: 0.01em;
    ">{label or input_type}</label>
''', unsafe_allow_html=True)

        if input_type == 'Text':
            is_multiline = element.get('isMultiline', False)
            if is_multiline:
                st.text_area(label or "Text input", disabled=True, key=f"input_{input_id}_{depth}", label_visibility="collapsed")
            else:
                st.text_input(label or "Text input", disabled=True, key=f"input_{input_id}_{depth}", label_visibility="collapsed")
        elif input_type == 'Number':
            st.number_input(label or "Number input", disabled=True, key=f"input_{input_id}_{depth}", label_visibility="collapsed")
        elif input_type == 'Date':
            st.date_input(label or "Date input", disabled=True, key=f"input_{input_id}_{depth}", label_visibility="collapsed")
        elif input_type == 'Time':
            st.time_input(label or "Time input", disabled=True, key=f"input_{input_id}_{depth}", label_visibility="collapsed")
        elif input_type == 'Toggle':
            st.checkbox(label or "Toggle", disabled=True, key=f"input_{input_id}_{depth}", label_visibility="collapsed")
        elif input_type == 'ChoiceSet':
            choices = element.get('choices', [])
            choice_titles = [c.get('title', '') for c in choices]
            st.selectbox(label or "Choice", choice_titles, disabled=True, key=f"input_{input_id}_{depth}", label_visibility="collapsed")

        st.markdown('</div>', unsafe_allow_html=True)

    elif elem_type == 'ActionSet':
        actions = element.get('actions', [])
        horizontal_alignment = element.get('horizontalAlignment') or 'left'

        if actions:
            # Create button layout with actual link functionality
            align_style = ""
            if horizontal_alignment and horizontal_alignment.lower() == 'center':
                align_style = "text-align: center;"
            elif horizontal_alignment and horizontal_alignment.lower() == 'right':
                align_style = "text-align: right;"

            st.markdown(f'<div style="margin: 1rem 0; {align_style}">', unsafe_allow_html=True)
            button_html = ""
            for action in actions:
                title = action.get('title', '')
                action_type = action.get('type', '')

                if title:
                    if action_type == 'Action.OpenUrl':
                        url = action.get('url', '')
                        if url:
                            button_html += f'''<a href="{url}" target="_blank" style="text-decoration: none;">
                <button style="
                    margin: 0.375rem;
                    padding: 0.75rem 1.5rem;
                    border: 2px solid #1D4ED8;
                    border-radius: 8px;
                    background: linear-gradient(135deg, #FFFFFF 0%, #F5F5F4 100%);
                    color: #1D4ED8;
                    cursor: pointer;
                    font-weight: 600;
                    font-size: 0.875rem;
                    font-family: 'Manrope', sans-serif;
                    transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
                    box-shadow: 0 2px 4px -1px rgba(0, 0, 0, 0.06);
                    letter-spacing: 0.01em;
                " onmouseover="this.style.background='linear-gradient(135deg, #DBEAFE 0%, #EFF6FF 100%)'; this.style.boxShadow='0 4px 6px -1px rgba(0, 0, 0, 0.1)'; this.style.transform='translateY(-1px)'"
                   onmouseout="this.style.background='linear-gradient(135deg, #FFFFFF 0%, #F5F5F4 100%)'; this.style.boxShadow='0 2px 4px -1px rgba(0, 0, 0, 0.06)'; this.style.transform='translateY(0)'"
                >{title}</button>
            </a>'''
                        else:
                            button_html += f'''<button style="
                    margin: 0.375rem;
                    padding: 0.75rem 1.5rem;
                    border: 1.5px solid #D6D3D1;
                    border-radius: 8px;
                    background: linear-gradient(135deg, #F5F5F4 0%, #E7E5E4 100%);
                    color: #78716C;
                    font-weight: 600;
                    font-size: 0.875rem;
                    font-family: 'Manrope', sans-serif;
                    letter-spacing: 0.01em;
                ">{title}</button>'''
                    else:
                        # Submit, Execute, etc. - show as disabled button
                        button_html += f'''<button style="
                    margin: 0.375rem;
                    padding: 0.75rem 1.5rem;
                    border: 1.5px solid #D6D3D1;
                    border-radius: 8px;
                    background: #E7E5E4;
                    color: #A8A29E;
                    font-weight: 600;
                    font-size: 0.875rem;
                    font-family: 'Manrope', sans-serif;
                    cursor: not-allowed;
                    letter-spacing: 0.01em;
                    opacity: 0.6;
                ">{title}</button>'''
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

# Contemporary Editorial Styling - Distinctive & Refined
st.markdown("""
<style>
    /* Import distinctive typography */
    @import url('https://fonts.googleapis.com/css2?family=Crimson+Pro:wght@400;600;700&family=Manrope:wght@400;500;600;700&display=swap');

    /* CSS Variables - Contemporary Editorial Design System */
    :root {
        /* Primary palette with depth */
        --primary-blue: #1D4ED8;
        --primary-blue-dark: #1E40AF;
        --primary-blue-light: #DBEAFE;
        --primary-blue-vivid: #3B82F6;
        --accent-amber: #F59E0B;
        --accent-emerald: #10B981;

        /* Sophisticated neutrals */
        --gray-50: #FAFAF9;
        --gray-100: #F5F5F4;
        --gray-200: #E7E5E4;
        --gray-300: #D6D3D1;
        --gray-400: #A8A29E;
        --gray-500: #78716C;
        --gray-600: #57534E;
        --gray-700: #44403C;
        --gray-800: #292524;
        --gray-900: #1C1917;
        --warm-white: #FFFEFB;

        /* Sophisticated shadows with warmth */
        --shadow-xs: 0 1px 2px 0 rgba(0, 0, 0, 0.03);
        --shadow-sm: 0 2px 4px -1px rgba(0, 0, 0, 0.06), 0 1px 2px -1px rgba(0, 0, 0, 0.04);
        --shadow-md: 0 4px 6px -1px rgba(0, 0, 0, 0.08), 0 2px 4px -2px rgba(0, 0, 0, 0.04);
        --shadow-lg: 0 10px 15px -3px rgba(0, 0, 0, 0.1), 0 4px 6px -4px rgba(0, 0, 0, 0.05);
        --shadow-xl: 0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 8px 10px -6px rgba(0, 0, 0, 0.04);

        /* Typography scale */
        --font-display: 'Crimson Pro', Georgia, serif;
        --font-body: 'Manrope', system-ui, -apple-system, sans-serif;
    }

    /* Smooth fade-in animation */
    @keyframes fadeInUp {
        from {
            opacity: 0;
            transform: translateY(10px);
        }
        to {
            opacity: 1;
            transform: translateY(0);
        }
    }

    @keyframes pulse {
        0%, 100% {
            opacity: 1;
        }
        50% {
            opacity: 0.5;
        }
    }

    @keyframes shimmer {
        0% {
            background-position: -1000px 0;
        }
        100% {
            background-position: 1000px 0;
        }
    }

    /* Hide Streamlit header and adjust layout */
    header {visibility: hidden;}
    .block-container {
        padding-top: 2.5rem;
        padding-bottom: 3rem;
        max-width: 900px;
    }

    /* Main app background with subtle texture */
    .main {
        background: linear-gradient(135deg, var(--warm-white) 0%, var(--gray-50) 100%);
        font-family: var(--font-body);
    }

    /* Title styling - Editorial serif */
    h1 {
        font-family: var(--font-display) !important;
        font-weight: 700 !important;
        color: var(--gray-900) !important;
        margin-bottom: 2rem !important;
        font-size: 2.5rem !important;
        letter-spacing: -0.025em !important;
        line-height: 1.2 !important;
        animation: fadeInUp 0.6s ease-out;
    }

    /* Subtle decorative element under title */
    h1::after {
        content: '';
        display: block;
        width: 60px;
        height: 3px;
        background: linear-gradient(90deg, var(--primary-blue) 0%, var(--accent-amber) 100%);
        margin-top: 1rem;
        border-radius: 2px;
    }

    /* User message styling - Clean with depth */
    .stChatMessage[data-testid="user-message"] {
        background: linear-gradient(135deg, var(--warm-white) 0%, #FFFFFF 100%) !important;
        border: 1px solid var(--gray-200) !important;
        border-radius: 16px !important;
        padding: 1.25rem 1.5rem !important;
        margin-bottom: 1.5rem !important;
        box-shadow: var(--shadow-md) !important;
        animation: fadeInUp 0.4s ease-out;
        position: relative;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    }

    .stChatMessage[data-testid="user-message"]::before {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        height: 2px;
        background: linear-gradient(90deg, transparent 0%, var(--gray-300) 50%, transparent 100%);
        border-radius: 16px 16px 0 0;
    }

    .stChatMessage[data-testid="user-message"]:hover {
        box-shadow: var(--shadow-lg) !important;
        transform: translateY(-1px);
    }

    .stChatMessage[data-testid="user-message"] p {
        color: var(--gray-800) !important;
        font-size: 1rem !important;
        line-height: 1.65 !important;
        margin: 0 !important;
        font-family: var(--font-body) !important;
    }

    /* Assistant message styling - Editorial blue with gradient accent */
    .stChatMessage[data-testid="assistant-message"] {
        background: linear-gradient(135deg, #EFF6FF 0%, var(--primary-blue-light) 100%) !important;
        border: 1.5px solid var(--primary-blue-vivid) !important;
        border-left: 4px solid var(--primary-blue) !important;
        border-radius: 16px !important;
        padding: 1.25rem 1.5rem !important;
        margin-bottom: 1.5rem !important;
        box-shadow: var(--shadow-lg) !important;
        animation: fadeInUp 0.5s ease-out;
        position: relative;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    }

    .stChatMessage[data-testid="assistant-message"]::after {
        content: '';
        position: absolute;
        top: -1px;
        left: -1px;
        right: -1px;
        bottom: -1px;
        background: linear-gradient(135deg, var(--primary-blue-vivid) 0%, transparent 30%);
        border-radius: 16px;
        opacity: 0.05;
        pointer-events: none;
    }

    .stChatMessage[data-testid="assistant-message"]:hover {
        box-shadow: var(--shadow-xl) !important;
        transform: translateY(-1px);
    }

    .stChatMessage[data-testid="assistant-message"] p {
        color: var(--gray-800) !important;
        font-size: 1rem !important;
        line-height: 1.7 !important;
        margin: 0 !important;
        font-family: var(--font-body) !important;
    }

    /* Chat input styling - Refined with focus state */
    .stChatInput {
        border: 2px solid var(--gray-300) !important;
        border-radius: 12px !important;
        background: linear-gradient(135deg, var(--warm-white) 0%, #FFFFFF 100%) !important;
        box-shadow: var(--shadow-sm) !important;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
    }

    .stChatInput:focus-within {
        border-color: var(--primary-blue-vivid) !important;
        box-shadow: 0 0 0 4px var(--primary-blue-light), var(--shadow-md) !important;
        transform: translateY(-1px);
    }

    .stChatInput textarea {
        font-size: 1rem !important;
        line-height: 1.6 !important;
        color: var(--gray-800) !important;
        font-family: var(--font-body) !important;
    }

    .stChatInput textarea::placeholder {
        color: var(--gray-400) !important;
        font-style: italic;
    }

    /* Button styling - New Chat button with gradient hover */
    .stButton > button {
        background: linear-gradient(135deg, var(--gray-100) 0%, var(--gray-50) 100%) !important;
        color: var(--gray-700) !important;
        border: 1.5px solid var(--gray-300) !important;
        border-radius: 10px !important;
        padding: 0.625rem 1.25rem !important;
        font-weight: 600 !important;
        font-size: 0.875rem !important;
        font-family: var(--font-body) !important;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
        box-shadow: var(--shadow-sm) !important;
        letter-spacing: 0.01em !important;
    }

    .stButton > button:hover {
        background: linear-gradient(135deg, var(--gray-200) 0%, var(--gray-100) 100%) !important;
        border-color: var(--primary-blue) !important;
        color: var(--primary-blue-dark) !important;
        box-shadow: var(--shadow-md) !important;
        transform: translateY(-1px);
    }

    .stButton > button:active {
        transform: translateY(0);
    }

    /* Status/thinking container - Subtle animation */
    .stStatus {
        background: linear-gradient(135deg, var(--warm-white) 0%, #FFFFFF 100%) !important;
        border: 1.5px solid var(--gray-200) !important;
        border-radius: 12px !important;
        box-shadow: var(--shadow-md) !important;
        animation: fadeInUp 0.4s ease-out;
    }

    .stStatus > div {
        color: var(--gray-700) !important;
        font-size: 0.875rem !important;
        font-weight: 600 !important;
        font-family: var(--font-body) !important;
    }

    /* Animated thinking indicator */
    .stStatus[data-state="running"]::before {
        content: '';
        display: inline-block;
        width: 8px;
        height: 8px;
        background: var(--primary-blue);
        border-radius: 50%;
        margin-right: 8px;
        animation: pulse 1.5s ease-in-out infinite;
    }

    /* Expander (for Adaptive Cards and JSON) - Refined */
    .streamlit-expanderHeader {
        background: linear-gradient(135deg, var(--gray-50) 0%, var(--warm-white) 100%) !important;
        border: 1.5px solid var(--gray-200) !important;
        border-radius: 10px !important;
        padding: 0.875rem 1.25rem !important;
        font-weight: 600 !important;
        font-size: 0.875rem !important;
        color: var(--gray-700) !important;
        font-family: var(--font-body) !important;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
    }

    .streamlit-expanderHeader:hover {
        background: linear-gradient(135deg, var(--gray-100) 0%, var(--gray-50) 100%) !important;
        border-color: var(--primary-blue-vivid) !important;
        color: var(--primary-blue-dark) !important;
        box-shadow: var(--shadow-sm) !important;
    }

    .streamlit-expanderContent {
        border: 1.5px solid var(--gray-200) !important;
        border-top: none !important;
        background: linear-gradient(180deg, var(--warm-white) 0%, #FFFFFF 100%) !important;
        padding: 1.25rem !important;
        border-radius: 0 0 10px 10px !important;
    }

    /* Caption text (status messages, suggestions) */
    .caption {
        color: var(--gray-500) !important;
        font-size: 0.8125rem !important;
        font-style: italic !important;
        font-family: var(--font-body) !important;
    }

    /* Code blocks - Monospace with refined styling */
    code {
        background: linear-gradient(135deg, var(--gray-100) 0%, var(--gray-50) 100%) !important;
        color: var(--gray-800) !important;
        padding: 0.2rem 0.4rem !important;
        border-radius: 4px !important;
        font-size: 0.875rem !important;
        border: 1px solid var(--gray-200) !important;
        font-family: 'SF Mono', Monaco, Consolas, monospace !important;
    }

    pre {
        background: linear-gradient(135deg, var(--gray-900) 0%, #1a1816 100%) !important;
        color: var(--gray-100) !important;
        padding: 1.25rem !important;
        border-radius: 10px !important;
        overflow-x: auto !important;
        box-shadow: var(--shadow-lg) !important;
        border: 1px solid var(--gray-800) !important;
    }

    pre code {
        background: transparent !important;
        border: none !important;
        padding: 0 !important;
        color: var(--gray-100) !important;
    }

    /* Info/Warning/Error messages - Enhanced with depth */
    .stAlert {
        border-radius: 12px !important;
        border-left-width: 4px !important;
        padding: 1.25rem 1.5rem !important;
        box-shadow: var(--shadow-md) !important;
        animation: fadeInUp 0.4s ease-out;
    }

    /* Spinner - Refined animation */
    .stSpinner > div {
        border-top-color: var(--primary-blue-vivid) !important;
        border-right-color: var(--primary-blue-light) !important;
    }

    /* Citation links - Distinctive with hover effect */
    a[target="_blank"] sup {
        color: var(--primary-blue) !important;
        font-weight: 700 !important;
        text-decoration: none !important;
        background: var(--primary-blue-light);
        padding: 0.125rem 0.25rem;
        border-radius: 3px;
        transition: all 0.2s ease;
    }

    a[target="_blank"]:hover sup {
        color: var(--primary-blue-dark) !important;
        background: var(--primary-blue-vivid) !important;
        color: white !important;
        transform: translateY(-1px);
        box-shadow: var(--shadow-sm);
    }
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
        if st.button("🔄 New Chat", help="Start a new conversation"):
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
                                    st.markdown('''<div style="
    border: 2px solid #E7E5E4;
    border-radius: 14px;
    padding: 1.75rem;
    background: linear-gradient(135deg, #FFFFFF 0%, #FAFAF9 100%);
    box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1), 0 4px 6px -4px rgba(0, 0, 0, 0.05);
    margin: 1rem 0;
    position: relative;
    animation: fadeInUp 0.5s ease-out;
">
<div style="
    position: absolute;
    top: 0;
    left: 0;
    right: 0;
    height: 3px;
    background: linear-gradient(90deg, #1D4ED8 0%, #3B82F6 50%, #10B981 100%);
    border-radius: 14px 14px 0 0;
"></div>
''', unsafe_allow_html=True)

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
                            status_placeholder.markdown(f'''<p class="caption" style="
    color: #78716C;
    font-size: 0.875rem;
    font-style: italic;
    margin: 0.75rem 0;
    font-family: 'Crimson Pro', serif;
    animation: pulse 2s ease-in-out infinite;
">💭 {msg_content}</p>''', unsafe_allow_html=True)
                        elif msg_type == 'thought':
                            # Collect reasoning/chain-of-thought
                            thoughts.append(msg_content)
                            # Update thinking display
                            with thinking_container.status("💡 Thinking...", expanded=False) as status:
                                for t in thoughts:
                                    task = t.get('task', 'Processing')
                                    text = t.get('text', '')
                                    st.markdown(f'''<div style="
    margin: 0.75rem 0;
    padding: 0.875rem 1rem;
    background: linear-gradient(135deg, #EFF6FF 0%, #DBEAFE 100%);
    border-left: 4px solid #1D4ED8;
    border-radius: 8px;
    box-shadow: 0 2px 4px -1px rgba(0, 0, 0, 0.06);
">
    <strong style="
        color: #1E40AF;
        font-size: 0.875rem;
        font-family: 'Manrope', sans-serif;
        font-weight: 700;
        letter-spacing: 0.01em;
    ">{task}</strong>
    <p style="
        color: #44403C;
        font-size: 0.875rem;
        margin: 0.5rem 0 0 0;
        line-height: 1.65;
        font-family: 'Manrope', sans-serif;
    ">{text}</p>
</div>''', unsafe_allow_html=True)
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
                        with thinking_container.status("💡 Reasoning", expanded=False, state="complete") as status:
                            for t in thoughts:
                                task = t.get('task', 'Processing')
                                text = t.get('text', '')
                                st.markdown(f'''<div style="
    margin: 0.75rem 0;
    padding: 0.875rem 1rem;
    background: linear-gradient(135deg, #ECFDF5 0%, #D1FAE5 100%);
    border-left: 4px solid #10B981;
    border-radius: 8px;
    box-shadow: 0 2px 4px -1px rgba(0, 0, 0, 0.06);
">
    <strong style="
        color: #065F46;
        font-size: 0.875rem;
        font-family: 'Manrope', sans-serif;
        font-weight: 700;
        letter-spacing: 0.01em;
    ">{task}</strong>
    <p style="
        color: #44403C;
        font-size: 0.875rem;
        margin: 0.5rem 0 0 0;
        line-height: 1.65;
        font-family: 'Manrope', sans-serif;
    ">{text}</p>
</div>''', unsafe_allow_html=True)

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
                st.markdown(f'''<div style="
    margin-top: 1.5rem;
    padding: 1.25rem 1.5rem;
    background: linear-gradient(135deg, #FEF3C7 0%, #FDE68A 100%);
    border-left: 4px solid #F59E0B;
    border-radius: 12px;
    box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.08);
    animation: fadeInUp 0.5s ease-out;
">
    <strong style="
        color: #92400E;
        font-size: 0.9375rem;
        font-family: 'Manrope', sans-serif;
        font-weight: 700;
        letter-spacing: 0.01em;
    ">💡 Suggestions:</strong>
    <span style="
        color: #78350F;
        font-size: 0.9375rem;
        margin-left: 0.5rem;
        font-family: 'Manrope', sans-serif;
        line-height: 1.65;
    ">{suggestions}</span>
</div>''', unsafe_allow_html=True)

        # Store response with HTML citations and adaptive cards for history
        st.session_state.messages.append({
            "role": "assistant",
            "content": response,
            "adaptive_cards": adaptive_cards if adaptive_cards else []
        })


if __name__ == "__main__":
    main()
