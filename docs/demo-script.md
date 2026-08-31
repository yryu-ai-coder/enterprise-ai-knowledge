# Demo Script

## Scenario

A user opens a SharePoint project page and asks the embedded assistant about project information. The reusable SPFx web part calls Azure Function Apps, which orchestrates Azure OpenAI or Microsoft Foundry.

## Demo Steps

1. Open SharePoint page containing the NextCore Projects AI Assistant web part.
2. Ask: `Summarize this project site for me.`
3. Confirm the web part sends request to Function App.
4. Confirm the backend returns answer and citations.
5. Explain that the web part is reusable across departments/projects.
6. Show that backend provider can switch from Azure OpenAI to Foundry without rewriting the web part.
