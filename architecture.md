# Architecture - NextCore Projects AI Assistant

## Purpose

Provide a reusable SharePoint AI Assistant web part for NextCore project sites. The web part remains generic and configurable. Azure Function Apps handle secure backend orchestration and AI/provider integration.

## Core Components

### 1. SPFx Web Part

Responsibilities:

- Render chat/search assistant UI inside SharePoint pages.
- Read SharePoint page/site/user context.
- Send user question and context to the backend API.
- Display answer, citations, errors, and loading status.
- Keep configuration in the web part property pane.

The web part must not contain Azure OpenAI keys, Foundry credentials, business secrets, or sensitive orchestration logic.

### 2. Azure Function App

Responsibilities:

- Expose `/api/chat` endpoint.
- Validate request body and caller context.
- Build system/user prompt.
- Call Azure OpenAI directly or delegate to Microsoft Foundry Agent.
- Optionally retrieve SharePoint/Graph/Search grounding context.
- Return a normalized answer/citation response to the web part.

### 3. AI Provider Layer

Initial path:

```text
Function App -> Azure OpenAI chat completions
```

Later path:

```text
Function App -> Microsoft Foundry Agent endpoint
```

The web part should not need to change when switching providers.

### 4. Knowledge/Data Layer

Possible grounding sources:

- Current SharePoint site documents
- SharePoint lists
- Microsoft Graph
- Azure AI Search index
- Foundry Agent knowledge/tools

## Recommended POC Scope

Start with a thin vertical slice:

```text
SPFx Web Part -> Function App -> Azure OpenAI -> Answer displayed in SharePoint
```

Then add grounding and citations.
