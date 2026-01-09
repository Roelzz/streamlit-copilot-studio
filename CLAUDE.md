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
The app automatically writes all SDK activities to `/tmp/activities_debug.json` during message exchanges. This file contains:
- Full activity types received from Copilot Studio
- Channel data and stream types
- Entity information (citations, claims)
- Search results and observations
- Chain-of-thought reasoning data

```bash
# Watch debug file in real-time
tail -f /tmp/activities_debug.json

# Pretty-print the debug file
cat /tmp/activities_debug.json | python -m json.tool

# Search for specific activity types
cat /tmp/activities_debug.json | python -m json.tool | grep -A 5 "type"
```

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
- Yields typed tuples during message processing: `('status', text)`, `('thought', dict)`, `('search_result', dict)`, `('content', chunk)`, `('final_content', text)`, `('citations', dict)`, `('suggestion', text)`
- Implements citation parsing and rendering (Unicode markers to numbered references with optional HTML links)
- Writes debug JSON to `/tmp/activities_debug.json` for troubleshooting SDK activity responses

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

## Key Dependencies

- `streamlit` - Web UI framework
- `streamlit-msal` - Azure authentication component (third-party library, not official Microsoft)
- `msal` - Microsoft Authentication Library
- `microsoft-agents-copilotstudio-client` - Copilot Studio SDK
- `microsoft-agents-activity` - Activity protocol types
- `aiohttp` - Async HTTP (SDK dependency)
