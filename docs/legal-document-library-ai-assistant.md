# Legal Document Library AI Assistant - POC Design

## Direction

Start with a legal/litigation document library assistant and keep the backend contract reusable for additional SharePoint surfaces.

```text
SharePoint Document Library
  -> SPFx contextual AI panel
  -> Reusable Function API
  -> Mock provider now
  -> Azure OpenAI / Microsoft Foundry later
  -> Microsoft Graph or Azure AI Search for document grounding later
```

## Why start with Document Library

A document library is the best first slice because the user's real target workflow is document-heavy legal/litigation material. It also creates a strong portfolio story:

```text
Built a contextual SharePoint AI assistant that analyzes selected legal documents from a document library and routes questions through a reusable backend API designed to later support SharePoint Lists and Site Pages.
```

## First-slice capabilities

The POC currently models these actions:

- Summarize selected litigation documents.
- Extract key dates and deadlines.
- Identify parties and claims.
- Find missing evidence or open questions.
- Draft a case timeline.

## Future SharePoint List expansion

The same API envelope supports list-based scenarios such as:

- Legal matter tracker.
- Deadline tracker.
- Evidence log.
- Discovery request tracker.
- Task/action item list.

The web part sends `contextType: "sharepoint-list"`, `listId`, `listTitle`, and `selectedItems`.

## Future Site Pages expansion

The same API envelope supports page-based scenarios such as:

- Legal brief page summarization.
- Matter summary pages.
- Internal knowledge pages.
- FAQ generation.
- Executive summary generation.

The web part sends `contextType: "site-page"`, `pageUrl`, and `pageTitle`. Later the backend retrieves page canvas content.

## Security and legal notes

- Treat real litigation documents as confidential and potentially privileged.
- Do not place document text, access tokens, or AI secrets directly in SPFx client code.
- Use mock data until an approved Azure tenant, model deployment, storage/indexing plan, and access-control design are ready.
- Use Azure Function App settings or Key Vault for secrets.
- Use Microsoft Graph delegated permissions or Azure AI Search index security trimming for real document retrieval.
- The assistant should be positioned as workflow/document review support, not legal advice.

## Implementation phases

### Phase 1 - Current local POC

- SPFx UI focused on Legal Document Library Assistant.
- Mock selected legal documents.
- Reusable API payload contains `scenario`, `mode`, `contextType`, `selectedFiles`, `selectedItems`, `pageUrl`.
- Mock backend returns legal-specific response, citations, and suggested actions.

### Phase 2 - Real SharePoint document metadata

- Replace hardcoded sample files with real SharePoint document library metadata.
- Read library/list/page context from SPFx page context or SharePoint REST/Graph.
- Let the user select documents inside the web part or build a ListView Command Set later.

### Phase 3 - Real document grounding

- Backend retrieves document content through Graph or indexes documents in Azure AI Search.
- Azure OpenAI / Foundry answers with citations.
- Add audit logging and request tracing.

### Phase 4 - Productionization

- Deploy backend as Azure Function App.
- Secure endpoint with Entra ID.
- Configure allowed origins.
- Move secrets to Key Vault or Function App settings.
- Upload `.sppkg` to App Catalog and deploy to target sites.
