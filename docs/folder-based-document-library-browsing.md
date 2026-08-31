# Folder-Based Document Library Browsing

## Current update

The Legal Document Library AI Assistant now groups real SharePoint document library files by folder using `FileDirRef`.

This replaces the earlier flat card layout, which became hard to use when the library had many files.

## UX behavior

The Document Library panel now shows:

```text
Folders panel        Files in selected folder
-------------        ------------------------
Matter A             Exhibit E-1.pdf
Matter B             Medical Records.pdf
Library root         Discovery Response.docx
```

User flow:

1. Select a folder.
2. The file list updates to files in that folder only.
3. Select a file.
4. Click one of the legal actions or ask a custom question.
5. The backend receives selected file metadata plus `folderPath`.

## SharePoint fields used

The SPFx web part calls:

```text
/_api/web/lists/getByTitle('<library>')/items
```

Selected fields:

- `Id`
- `FileLeafRef`
- `FileRef`
- `FileDirRef`
- `File_x0020_Type`
- `Modified`
- `FSObjType`

The query filters to files only:

```text
FSObjType eq 0
```

The current top limit is `200` files for the POC.

## Backend payload addition

The Ask AI payload now includes:

```json
{
  "folderPath": "/sites/enterprise-ai-knowledge/Litigation Documents/Matter A",
  "selectedFiles": [
    {
      "name": "Exhibit E-1.pdf",
      "serverRelativeUrl": "/sites/enterprise-ai-knowledge/Litigation Documents/Matter A/Exhibit E-1.pdf",
      "folderPath": "/sites/enterprise-ai-knowledge/Litigation Documents/Matter A",
      "displayFolderPath": "Matter A"
    }
  ]
}
```

## Next improvements

- Add folder search/filter.
- Add multi-file selection per folder.
- Add "Select all files in folder".
- Add breadcrumb display for deeply nested folders.
- Add pagination or Microsoft Graph delta query for very large libraries.
