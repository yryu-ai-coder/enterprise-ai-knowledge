# Ask AI Popup and Parent/Child Folder Browsing

## Changes

The Legal Document Library AI Assistant UI was simplified again after feedback.

### Ask AI command

The previous `AI legal assistant` toggle in the command bar was removed. It is now replaced with a clearer command button:

```text
Ask AI
```

Clicking `Ask AI` opens the legal AI assistant as a popup side panel. This keeps the document library area focused on folder/file browsing and avoids permanently consuming screen width.

### Parent/child folder browsing

The Litigation Documents view now starts by showing only parent folders. The flow is:

```text
Parent folders
  -> click parent folder
    -> child folders under that parent
      -> click child folder
        -> files in that child folder
```

This avoids showing every folder path at once and makes the experience closer to a normal SharePoint document-library browsing pattern.

## Verified

SPFx build and package generation succeeded:

```powershell
npm run build
```

Package regenerated at:

```text
spfx-webpart/sharepoint/solution/nextcore-ai-knowledge-webpart.sppkg
```
