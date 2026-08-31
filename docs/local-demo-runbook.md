# Enterprise AI Knowledge Workspace - Local Demo Runbook

This runbook captures the current working POC flow for the NextCore SharePoint AI web part.

## Current milestone

The local foundation is working:

- SPFx React web part renders in SharePoint Online Workbench.
- The default SPFx welcome screen was replaced with the Enterprise AI Knowledge Workspace UI.
- A mock Function-style backend responds on `POST /api/chat` and `POST /api/ask`.
- The web part default endpoint is `http://localhost:7072/api/chat`.
- The production package is generated at `spfx-webpart/sharepoint/solution/nextcore-ai-knowledge-webpart.sppkg`.

## Terminal 1 - Start the mock Function backend

```powershell
cd D:\NextCore\Projects\SPFx\NextCore-projects\function-app
$env:PORT=7072
npm run start:mock
```

Expected output:

```text
NextCore Function App dev server running at http://localhost:7072
Available endpoints: GET /api/health, POST /api/chat, POST /api/ask
```

If you see `EADDRINUSE`, the port is already in use. Check whether it is already working:

```powershell
curl.exe http://localhost:7072/api/health
```

## Terminal 2 - Start the SPFx debug server

```powershell
cd D:\NextCore\Projects\SPFx\NextCore-projects\spfx-webpart
npm run start
```

Expected behavior:

- SPFx dev server listens on `https://localhost:4321`.
- Keep this terminal open while using SharePoint Workbench.

## Browser - Open SharePoint Workbench

Open the SharePoint site workbench:

```text
https://<tenant>.sharepoint.com/sites/ai-knowledge-workspace/_layouts/15/workbench.aspx
```

Allow debug scripts when prompted. Then add the web part:

```text
+ Add web part -> Enterprise AI Knowledge Workspace / AiKnowledgeWorkspace
```

## Expected UI

The page should show:

- Header: `Legal Document Library AI Assistant`
- Left navigation: Document Library, SharePoint List, Site Page
- Center: Legal document-library context with sample litigation documents
- Right: Ask Legal AI / Document Assistant panel
- Backend badge pointing to `http://localhost:7072/api/chat`
- Legal-specific action buttons: summarize documents, extract deadlines, identify parties/claims, find missing evidence, draft timeline

## End-to-end mock test

In the Ask AI panel, use:

```text
Summarize this workspace
```

Expected response:

```text
Mock response for: "Summarize this workspace"
This confirms the SPFx web part can call the Function App backend.
```

Expected citation:

```text
NextCore Projects Architecture
```

## Direct API verification

```powershell
curl.exe -X POST "http://localhost:7072/api/chat" -H "Content-Type: application/json" -d '{ "question": "Summarize this workspace", "siteUrl": "https://sharepoint.local" }'
```

Expected JSON includes:

```json
{
  "provider": "mock",
  "status": "success"
}
```

## Build and package

```powershell
cd D:\NextCore\Projects\SPFx\NextCore-projects\spfx-webpart
npm run build
```

Expected package:

```text
D:\NextCore\Projects\SPFx\NextCore-projects\spfx-webpart\sharepoint\solution\nextcore-ai-knowledge-webpart.sppkg
```

## Next production steps

1. Upload `.sppkg` to SharePoint App Catalog.
2. Add the web part to the real `Enterprise AI Knowledge Workspace` page.
3. Deploy the backend to Azure Function App.
4. Replace local endpoint with Azure Function URL.
5. Replace mock provider with Azure OpenAI or Microsoft Foundry provider.
6. Add SharePoint document grounding through Graph or Azure AI Search.

## Troubleshooting quick map

| Symptom | Meaning | Action |
|---|---|---|
| `func is not recognized` | Azure Functions Core Tools missing | Use `npm run start:mock` for local POC |
| `npm E401` while installing Core Tools | npm global auth/token issue | Use mock server now; fix npm login later |
| `EADDRINUSE :7072` | Backend already running on 7072 | Do not start another server; test `/api/health` |
| Browser shows `{ "error": "Not found" }` on `/api/chat` | `/api/chat` requires POST | Use Ask AI button or `curl.exe -X POST` |
| Old SPFx welcome screen remains | Browser cached old bundle | Restart `npm run start` and press `Ctrl+F5` |
