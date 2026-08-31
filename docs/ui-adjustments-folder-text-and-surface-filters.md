# UI Adjustment - Folder Text and Surface Filters

## Changes

The SharePoint-library-style UI was adjusted after usability feedback.

### Fixed folder text cut-off

Folder and file names now wrap instead of being clipped. The table grid was changed to give the name column flexible width and smaller metadata columns.

### Restored SharePoint surface filters

A compact tab filter was added below the command bar:

- Document Library
- SharePoint List
- Site Pages

Current behavior:

- `Document Library` is the active implemented mode.
- `SharePoint List` and `Site Pages` are restored as roadmap filters with placeholder messages, so the reusable backend direction remains visible without making the main UI complex.

## Verified

SPFx build succeeded:

```powershell
npm run build
```

Package regenerated:

```text
spfx-webpart/sharepoint/solution/nextcore-ai-knowledge-webpart.sppkg
```
