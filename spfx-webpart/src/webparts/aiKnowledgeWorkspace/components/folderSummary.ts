/** Builds the explicit, folder-scoped Ask AI prompt used by the summary action. */
export function buildFolderSummaryQuestion(folderName: string): string {
  return `Summarize the indexed documents in the ${folderName} folder and its subfolders. Include the document types and main topics, and use only confirmed source evidence.`;
}
