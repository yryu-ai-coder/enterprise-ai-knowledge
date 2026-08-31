# NextCore Projects - Reusable SharePoint AI Assistant Web Part

This project is a starter architecture for a reusable SharePoint Framework (SPFx) web part backed by Azure Function Apps and Azure OpenAI / Microsoft Foundry.

## Goal

Create a reusable SharePoint web part that can be placed on multiple SharePoint sites/pages while delegating secure AI orchestration, grounding, and business logic to Azure Function Apps.

## High-level Architecture

```text
SharePoint Page
  └─ SPFx Web Part (reusable UI)
        └─ Azure Function App API
              ├─ Request validation
              ├─ Auth / permission checks
              ├─ Prompt / orchestration layer
              ├─ Azure OpenAI direct call
              ├─ Microsoft Foundry Agent option
              └─ SharePoint / Graph / Search grounding
```

## Folders

```text
NextCore-projects/
├── architecture.md
├── api-contract.md
├── docs/
├── spfx-webpart/
└── function-app/
```

## Build Order

1. Confirm the web part can call the Function App hello endpoint.
2. Add Azure OpenAI direct chat response.
3. Add citations and SharePoint grounding.
4. Add Microsoft Foundry Agent backend option.
5. Harden authentication, logging, permissions, and deployment.

## Current Local Demo

The current local POC uses:

- `function-app`: mock Function-style backend on `http://localhost:7072`.
- `spfx-webpart`: SPFx React web part with Legal Document Library AI Assistant UI.
- SharePoint Workbench: debug host for validating the web part before App Catalog deployment.
- Reusable backend contract: starts with legal document libraries and is designed to expand to SharePoint Lists and Site Pages.

Design note:

```text
docs/legal-document-library-ai-assistant.md
```

Document library metadata integration:

```text
docs/real-document-library-metadata.md
```

Folder-based document browsing:

```text
docs/folder-based-document-library-browsing.md
```

Runbook:

```text
docs/local-demo-runbook.md
```

Real AI model provider setup:

```text
docs/real-ai-model-provider.md
```

Current package:

```text
spfx-webpart/sharepoint/solution/nextcore-ai-knowledge-webpart.sppkg
```
