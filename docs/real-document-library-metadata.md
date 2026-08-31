# Real SharePoint Document Library Metadata Integration

## Current status

The SPFx web part now reads file metadata from the configured SharePoint document library instead of using hardcoded sample legal document cards.

Default library name:

```text
Litigation Documents
```

Configurable web part property:

```text
Document library name
```

## What the web part reads

The current implementation calls SharePoint REST through SPFx `SPHttpClient`:

```text
/_api/web/lists/getByTitle('Litigation Documents')/items
```

Selected fields:

- `Id`
- `FileLeafRef` - file name
- `FileRef` - server-relative file URL
- `File_x0020_Type` - file extension/type
- `Modified` - modified date
- `FSObjType` - filters folders out and keeps files only

The web part currently reads the latest 25 files ordered by modified date.

## Current behavior

1. Page loads.
2. Web part checks backend health.
3. Web part loads files from `Litigation Documents`.
4. File cards appear in the center panel.
5. User selects a file.
6. Ask AI sends selected file metadata to the reusable backend API.
7. Backend mock response returns legal document analysis and citations.

## Backend payload now includes actual SharePoint file metadata

```json
{
  "scenario": "legal-document-library",
  "mode": "legal-analysis",
  "contextType": "document-library",
  "siteUrl": "https://tenant.sharepoint.com/sites/enterprise-ai-knowledge",
  "libraryName": "Litigation Documents",
  "selectedFiles": [
    {
      "id": 1,
      "name": "Sample Legal Memo.docx",
      "type": "DOCX",
      "libraryTitle": "Litigation Documents",
      "url": "https://tenant.sharepoint.com/sites/enterprise-ai-knowledge/Shared Documents/...",
      "serverRelativeUrl": "/sites/enterprise-ai-knowledge/...",
      "snippet": "SharePoint file from Litigation Documents. Last modified ...",
      "lastModified": "..."
    }
  ],
  "question": "Summarize selected litigation documents"
}
```

## Important limitation

This phase reads **file metadata only**, not full document contents. That is intentional for safety and scope.

To summarize real document contents, the next phase should retrieve/index document text server-side using one of these approaches:

1. Microsoft Graph file content download + server-side text extraction.
2. Azure AI Search index over the SharePoint document library.
3. SharePoint Syntex / Microsoft Search / Graph connectors depending on licensing and architecture.

## User test steps

1. In the SharePoint site, create a document library named:

```text
Litigation Documents
```

2. Upload 1-3 safe test documents. Start with sanitized/sample files, not privileged real case files.

3. Restart/refresh SPFx debug if needed:

```powershell
cd D:\NextCore\Projects\SPFx\NextCore-projects\spfx-webpart
npm run start
```

4. Refresh the SharePoint page with `Ctrl + F5`.

5. Confirm the center panel displays the uploaded file names.

6. Select a file and click `Ask AI`.

7. Confirm the response citation uses the selected real SharePoint file name and URL.

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| `Could not read library 'Litigation Documents'` | Library does not exist or name differs | Create library or update web part property `Document library name` |
| `no files found` | Library exists but has no files in root/query scope | Upload files and click Refresh files |
| Old sample cards still show | Old bundle/cache | Restart `npm run start`, refresh browser with `Ctrl + F5` |
| Ask AI disabled | No file selected / no file loaded | Upload files or fix library name |
| Backend offline | Mock backend not running | Start `function-app` with `$env:PORT=7072; npm run start:mock` |
