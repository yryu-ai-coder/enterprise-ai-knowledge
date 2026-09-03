import assert from 'node:assert/strict';
import test from 'node:test';
import { buildGroundedLibraryRequest } from '../src/services/ragAnswerService';

test('buildGroundedLibraryRequest converts retrieved chunks into bounded source-linked chat context', () => {
  const request = buildGroundedLibraryRequest(
    {
      question: 'Find Ava academic plan communications.',
      siteUrl: 'https://youngryu.sharepoint.com/sites/enterprise-ai-knowledge',
      libraryName: 'Litigation Documents'
    },
    [
      {
        id: 'chunk-one',
        content: 'Ava discussed the independent study plan and a September deadline.',
        documentName: 'consultation.pdf',
        documentUrl: 'https://youngryu.sharepoint.com/consultation.pdf',
        folderPath: '/Litigation Documents',
        fileType: 'pdf',
        lastModified: '2026-09-03T12:00:00Z',
        chunkOrdinal: 0
      }
    ]
  );

  assert.equal(request.knowledgeScope, 'library-wide-rag');
  assert.equal(request.selectedFiles?.[0].name, 'consultation.pdf');
  assert.match(request.selectedFiles?.[0].snippet || '', /September deadline/);
  assert.match(request.documentSnippets?.[0] || '', /Retrieved excerpt from consultation\.pdf/);
});
