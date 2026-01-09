# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

A Streamlit-based web chat interface for Microsoft Copilot Studio agents. The app uses Azure Entra ID authentication (via MSAL) to obtain access tokens and communicates with Copilot Studio using the M365 Agents SDK.

## Development Commands

Using uv (recommended):
```bash
# Sync dependencies (creates virtual environment and installs packages)
uv sync

# Run the application
uv run streamlit run app.py
```

Using pip:
```bash
# Install dependencies
pip install -r requirements.txt

# Run the application
streamlit run app.py
```

Access the app at http://localhost:8501

The app includes `watchdog` for automatic file watching - Streamlit will automatically reload when you modify Python files during development.

## Debugging and Logs

### Streamlit Logs
```bash
# Run with verbose logging
uv run streamlit run app.py --logger.level=debug

# Disable file watcher (useful if it's causing issues)
uv run streamlit run app.py --server.fileWatcherType=none

# View server logs (Streamlit outputs to stdout)
# Look for connection errors, module import issues, etc.
```

### Activity Debug JSON
When `DEBUG_MODE=true` in `.env`, the app writes all SDK activities to a debug JSON file. This file contains:
- Full activity types received from Copilot Studio
- Channel data and stream types
- Entity information (citations, claims)
- Search results and observations
- Chain-of-thought reasoning data

```bash
# Enable debug mode in .env
DEBUG_MODE=true

# Watch debug file in real-time (default location in system temp dir)
tail -f "$(python3 -c 'import tempfile; print(tempfile.gettempdir())')/activities_debug.json"

# Pretty-print the debug file
cat "$(python3 -c 'import tempfile; print(tempfile.gettempdir())')/activities_debug.json" | python -m json.tool

# Or set a custom debug file location
DEBUG_FILE=/path/to/your/debug.json
```

**Security Note:** Debug files are created with restrictive permissions (0o600 - owner read/write only) and are disabled by default in production.

### Common Issues

**Authentication failures:**
- Check `.env` configuration matches Azure Portal settings
- Verify app registration has SPA redirect URI: `http://localhost:8501`
- Ensure API permission `CopilotStudio.Copilots.Invoke` is granted and admin-consented

**Empty or no response from Copilot:**
- Check `/tmp/activities_debug.json` for activity types received
- Verify `COPILOT_ENVIRONMENT_ID` and `COPILOT_AGENT_IDENTIFIER` are correct
- Ensure the agent is published in Copilot Studio

**Citations not showing:**
- Inspect `activities_debug.json` for entities with type containing "Claim"
- Check that search results are being captured in event activities
- Verify citation IDs match between entities and inline markers

**Module import errors:**
- Run `uv sync` to ensure all dependencies are installed
- Check Python version is 3.9+: `python --version`

## Required Configuration

Copy `.env.example` to `.env` and configure:
- `COPILOT_ENVIRONMENT_ID` - Found in Copilot Studio > Settings > Advanced > Metadata
- `COPILOT_AGENT_IDENTIFIER` - Agent schema name from the same location
- `AZURE_TENANT_ID` - Your Azure tenant ID
- `AZURE_APP_CLIENT_ID` - Your app registration client ID (must have SPA redirect to `http://localhost:8501` and `CopilotStudio.Copilots.Invoke` API permission)

Optional configuration:
- `DEBUG_MODE` - Set to `true` to enable debug JSON output (default: `false`)
- `DEBUG_FILE` - Custom path for debug output (default: system temp directory)

The app validates all required environment variables at startup and shows helpful error messages if any are missing.

## Architecture

### Two-File Structure

**app.py** - Main Streamlit UI application
- Handles authentication via `streamlit-msal` library
- Manages session state for messages and client instance
- Orchestrates the chat UI with message display, streaming responses, and collapsible reasoning display
- Processes async responses from the client, including status updates, chain-of-thought, search results, content chunks, citations, and suggestions

**copilot_client.py** - Copilot Studio SDK wrapper
- `CopilotStudioClient` class wraps the M365 Agents SDK's `CopilotClient`
- Manages conversation lifecycle (start conversation, send messages)
- Yields typed tuples during message processing: `('status', text)`, `('thought', dict)`, `('search_result', dict)`, `('content', chunk)`, `('final_content', text)`, `('adaptive_card', dict)`, `('attachment', dict)`, `('citations', dict)`, `('suggestion', text)`
- Implements citation parsing and rendering (Unicode markers to numbered references with optional HTML links)
- Extracts and processes attachments including Adaptive Cards from activity responses
- Writes debug JSON for troubleshooting SDK activity responses (when DEBUG_MODE=true)

### Adaptive Card Support

Copilot Studio can send Adaptive Cards as attachments. The app:
- Detects attachments with contentType containing "adaptive"
- Displays cards in expandable sections (📋 Adaptive Card)
- Shows both the raw JSON structure and basic rendering of common elements
- Supports TextBlock (with bold formatting) and Image elements
- Stores cards in message history for persistence

To send adaptive cards from Copilot Studio, use the "Send a message" action with attachments.

### Citation Handling

Copilot Studio responses contain Unicode citation markers: `\ue200cite\ue202{citation_id}\ue201`

- `clean_citations()` parses these markers and converts them to numbered references `[1]`, `[2]`, etc.
- Citation metadata (URLs, titles) comes from two sources:
  1. `entities` in message activities (schema.org Claim objects)
  2. `search_result` data in event activities (matched by index)
- During streaming, citations display as plain text `[1]`
- After completion, they become clickable HTML links if URLs are available
- A references section is appended with all citation details

### Async Event Loop Management

Streamlit doesn't provide a natural async context. The app creates new event loops as needed:
```python
loop = asyncio.new_event_loop()
asyncio.set_event_loop(loop)
try:
    result = loop.run_until_complete(async_function())
finally:
    loop.close()
```

### Activity Types

The SDK yields various `ActivityTypes`:
- `ActivityTypes.event` - Contains chain-of-thought reasoning (`value.thought`) and search results (`value.observation.search_result`)
- `ActivityTypes.typing` - Streaming chunks (when `chunkType='delta'`) or status messages (when `streamType='informative'`)
- `ActivityTypes.message` - Final response with entities (citations) and suggested actions
- `ActivityTypes.end_of_conversation` - Signals conversation termination

### Session State Management

- `st.session_state.messages` - Chat history as list of `{"role": "user|assistant", "content": str}`
- `st.session_state.client` - `CopilotStudioClient` instance (persisted across reruns to maintain conversation_id)
- New conversation button clears both and triggers rerun

## Security Features

- **XSS Protection:** All HTML content is sanitized using `bleach` with a whitelist of safe tags before rendering
- **Error Handling:** Comprehensive try-catch blocks around all API calls with user-friendly error messages
- **Timeout Protection:** 30-second timeout for connections, 5-minute timeout for responses
- **Environment Validation:** Required configuration is validated at startup with clear error messages
- **Debug Security:** Debug output is disabled by default and uses restrictive file permissions when enabled

## Key Dependencies

- `streamlit` - Web UI framework
- `streamlit-msal` - Azure authentication component (third-party library, not official Microsoft)
- `msal` - Microsoft Authentication Library
- `microsoft-agents-copilotstudio-client` - Copilot Studio SDK
- `microsoft-agents-activity` - Activity protocol types
- `aiohttp` - Async HTTP (SDK dependency)
- `bleach` - HTML sanitization for XSS protection
- `watchdog` - File watching for hot reload
