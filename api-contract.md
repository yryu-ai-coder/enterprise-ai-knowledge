# API Contract

The backend API is intentionally reusable. The first production scenario is a **Legal Document Library AI Assistant**, but the same request envelope can later support SharePoint Lists and Site Pages.

## Current primary scenario

```text
Legal Document Library Assistant
```

Goal:

- User uploads litigation/legal documents to a SharePoint document library.
- SPFx Web Part sends selected document metadata and the user's question to the backend.
- Backend later uses Microsoft Graph, Azure AI Search, Azure OpenAI, or Microsoft Foundry to generate source-grounded responses.

> Important: Do not put legal documents, Azure OpenAI keys, tokens, or other secrets in the SPFx client bundle. The SPFx client sends context metadata; the backend performs secure document retrieval and AI orchestration.

## POST /api/chat

### Request - Legal Document Library

```json
{
  "question": "Summarize selected litigation documents",
  "scenario": "legal-document-library",
  "mode": "legal-analysis",
  "contextType": "document-library",
  "siteUrl": "https://tenant.sharepoint.com/sites/legal-workspace",
  "libraryName": "Litigation Documents",
  "folderPath": "/Shared Documents/Matter A",
  "selectedFiles": [
    {
      "name": "Complaint and Answer.pdf",
      "url": "https://tenant.sharepoint.com/sites/legal-workspace/Shared%20Documents/Matter%20A/Complaint%20and%20Answer.pdf",
      "uniqueId": "sharepoint-file-unique-id",
      "fileType": "pdf",
      "libraryTitle": "Litigation Documents",
      "lastModified": "2026-08-30T00:00:00Z",
      "snippet": "Optional client-side or indexed snippet. Real document extraction should happen server-side."
    }
  ],
  "documentSnippets": [],
  "pageContext": {
    "webTitle": "Enterprise AI Knowledge Workspace",
    "userEmail": "user@company.com"
  },
  "knowledgeScope": "legal-document-library"
}
```

### Request - Future SharePoint List mode

```json
{
  "question": "Find overdue legal tasks",
  "scenario": "legal-sharepoint-list",
  "mode": "list-analysis",
  "contextType": "sharepoint-list",
  "siteUrl": "https://tenant.sharepoint.com/sites/legal-workspace",
  "listId": "list-guid",
  "listTitle": "Legal Matter Tracker",
  "selectedItems": [
    {
      "id": 12,
      "title": "Discovery deadline review",
      "listTitle": "Legal Matter Tracker",
      "fields": {
        "Status": "Open",
        "Priority": "High",
        "DueDate": "2026-09-15"
      }
    }
  ]
}
```

### Request - Future Site Page mode

```json
{
  "question": "Create an executive brief from this page",
  "scenario": "legal-site-page",
  "mode": "page-summary",
  "contextType": "site-page",
  "siteUrl": "https://tenant.sharepoint.com/sites/legal-workspace",
  "pageUrl": "https://tenant.sharepoint.com/sites/legal-workspace/SitePages/Matter-A-Brief.aspx",
  "pageTitle": "Matter A Brief"
}
```

## Success Response

```json
{
  "answer": "Here is the grounded legal document summary...",
  "citations": [
    {
      "title": "Complaint and Answer.pdf",
      "url": "https://tenant.sharepoint.com/sites/legal-workspace/Shared%20Documents/Matter%20A/Complaint%20and%20Answer.pdf",
      "snippet": "Relevant supporting text."
    }
  ],
  "provider": "azure-openai",
  "requestId": "req-001",
  "status": "success",
  "suggestedActions": [
    "Extract key dates and deadlines",
    "Identify parties and claims",
    "Draft a case timeline"
  ],
  "metadata": {
    "scenario": "legal-document-library",
    "mode": "legal-analysis",
    "contextType": "document-library",
    "selectedFileCount": 1
  }
}
```

## Error Response

```json
{
  "answer": "",
  "citations": [],
  "provider": "mock",
  "requestId": "req-001",
  "status": "error",
  "error": "Question is required."
}
```

## Security Notes

- API keys must stay in Function App settings or Key Vault.
- Browser/client must never receive Azure OpenAI, Foundry, Graph, or SharePoint app-only secrets.
- Legal/litigation materials may contain privileged, confidential, or personally identifiable information. Do not send real documents to non-approved AI services.
- The production backend should enforce allowed origins, user auth, tenant/site scope, request size limits, logging, and audit boundaries.
- The first POC may use mock metadata; real document text should be retrieved server-side through Microsoft Graph or Azure AI Search.
