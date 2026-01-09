# Streamlit Copilot Studio Client

Web chat interface for Microsoft Copilot Studio agents.

## Prerequisites

- Python 3.9+
- A published Agent in [Copilot Studio](https://copilotstudio.microsoft.com)
- An Azure Entra ID app registration

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

## Run

```bash
pip install -r requirements.txt
streamlit run app.py
```

Open http://localhost:8501 and sign in with Microsoft.
