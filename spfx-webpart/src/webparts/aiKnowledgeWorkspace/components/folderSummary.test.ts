import { buildFolderSummaryQuestion } from './folderSummary';

describe('folder summary request', () => {
  it('asks for an evidence-grounded summary of the selected folder and its descendants', () => {
    expect(buildFolderSummaryQuestion('01_Legal_Official')).toBe(
      'Summarize the indexed documents in the 01_Legal_Official folder and its subfolders. Include the document types and main topics, and use only confirmed source evidence.'
    );
  });
});
