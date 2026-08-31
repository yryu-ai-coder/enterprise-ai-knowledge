# Deployment Notes

## SPFx Web Part

- Build/package with SPFx toolchain.
- Deploy package to SharePoint App Catalog.
- Add web part to target SharePoint pages.
- Configure backend API URL in property pane.

## Function App

- Deploy Node.js Azure Functions app.
- Configure environment variables.
- For mock-only local testing:
  - `AI_PROVIDER=mock`
- For real Azure OpenAI model calls:
  - `AI_PROVIDER=azure-openai`
  - `AZURE_OPENAI_ENDPOINT`
  - `AZURE_OPENAI_API_KEY`
  - `AZURE_OPENAI_DEPLOYMENT`
  - `AZURE_OPENAI_API_VERSION`
  - optional: `AI_TEMPERATURE`, `AI_MAX_TOKENS`
- For direct OpenAI API model calls:
  - `AI_PROVIDER=openai`
  - `OPENAI_API_KEY`
  - `OPENAI_MODEL`
  - optional: `OPENAI_BASE_URL`, `AI_TEMPERATURE`, `AI_MAX_TOKENS`
- For a Foundry Agent adapter:
  - `AI_PROVIDER=foundry`
  - `FOUNDRY_AGENT_ENDPOINT`
  - `FOUNDRY_AGENT_API_KEY`
- Configure CORS for SharePoint domain during POC.
- Never place model API keys in SPFx client code. Keep them server-side only.
- See `docs/real-ai-model-provider.md`.
