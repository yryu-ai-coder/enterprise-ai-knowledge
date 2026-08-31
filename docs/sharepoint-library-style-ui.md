# SharePoint-Library-Style UI Simplification

## Why changed

The previous folder/file split panel was functional but visually too heavy when many legal files were loaded. The updated design follows a simpler SharePoint document library view pattern.

## New layout

```text
Command bar
  Upload | New | Settings | AI legal assistant | Refresh

Main card
  문서 작업 영역
  Litigation Documents

  Folder rows OR file rows
  -------------------------------------------------
  selector | Name / Category | Files/Modified | Modified by

Right AI panel
  Ask Legal AI
  Selected context
  Suggested actions
  Question
  Response
  Citations
```

## Interaction

1. Initial view shows folders as table rows.
2. Selecting a folder opens the file list for that folder.
3. Back button returns to folders.
4. Selecting a file updates the AI panel selected context.
5. Ask AI sends selected file metadata and folder path to the backend.

## Design intent

The UI should feel like a SharePoint document library with an AI side panel, not a separate complex dashboard.

## Verified

`npm run build` succeeded and regenerated:

```text
spfx-webpart/sharepoint/solution/nextcore-ai-knowledge-webpart.sppkg
```
