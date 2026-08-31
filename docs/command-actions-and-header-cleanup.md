# Command Actions and Header Cleanup

## Changes

The command bar and document library header were updated after feedback.

### Implemented command actions

- `Upload` now opens a file picker and uploads selected file(s) to the current SharePoint folder using SharePoint REST:
  - selected child/file folder when available
  - selected parent folder when browsing a parent
  - library root when no folder is selected
- `New` now prompts for a folder name and creates a SharePoint folder under the current location.
- `Settings` now opens a popup settings panel showing:
  - SharePoint site URL
  - Document library name
  - backend API endpoint
  - current upload target
  - loaded metadata count
  - backend health recheck action

### Header cleanup

Removed:

- `문서 작업 영역`
- `Enterprise AI Knowledge Workspace · 54 file(s) · 8 folder(s)` style subtitle

Moved the backend health indicator to the former `문서 작업 영역` location above `Litigation Documents`.

### Typography

Updated the web part font stack to a more polished display style:

```text
Aptos Display, Segoe UI Variable Display, Trebuchet MS, Segoe UI, Arial, sans-serif
```

## Verified

SPFx build and package generation succeeded without warnings:

```powershell
npm run build
```

Package regenerated at:

```text
spfx-webpart/sharepoint/solution/nextcore-ai-knowledge-webpart.sppkg
```
