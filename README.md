# Streamlit Copilot Studio Client

Web chat interface for Microsoft Copilot Studio agents using Azure Entra ID authentication and M365 Agents SDK.

## Quick Start

**Prerequisites:** Python 3.9+, UV package manager, published Copilot Studio agent, Azure app registration

```bash
# Install dependencies
uv sync

# Run the app (IMPORTANT: use 'streamlit run', not just 'uv run app.py')
uv run streamlit run app.py
```

Open http://localhost:8501 and sign in with Microsoft.

## Azure App Setup

1. Create an app registration in [Azure Portal](https://portal.azure.com)
2. Add a **Single-page application (SPA)** redirect URI: `http://localhost:8501`
3. Add API permission: **Power Platform API** > **CopilotStudio.Copilots.Invoke**

## Configuration

```bash
cp .env.example .env
```

Edit `.env` with your values:
- `COPILOT_ENVIRONMENT_ID` - From Copilot Studio > Settings > Advanced > Metadata
- `COPILOT_AGENT_IDENTIFIER` - Agent schema name from the same location
- `AZURE_TENANT_ID` - Your Azure tenant ID
- `AZURE_APP_CLIENT_ID` - Your app registration client ID

Optional:
- `DEBUG_MODE` - Set to `true` to write activity debug JSON (default: `false`)
- `DEBUG_FILE` - Custom path for debug output (default: system temp directory)

## Development

```bash
# Run with debug logging
uv run streamlit run app.py --logger.level=debug

# View debug activities (when DEBUG_MODE=true in .env)
tail -f "$(python3 -c 'import tempfile; print(tempfile.gettempdir())')/activities_debug.json"
```

See CLAUDE.md for full troubleshooting guide, architecture details, and common issues.
