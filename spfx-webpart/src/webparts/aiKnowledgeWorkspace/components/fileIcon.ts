export type FileIconKind = 'pdf' | 'word' | 'excel' | 'powerpoint' | 'generic';

export function getFileIconKind(fileType: string): FileIconKind {
  switch (fileType.trim().toLowerCase()) {
    case 'pdf': return 'pdf';
    case 'doc':
    case 'docx': return 'word';
    case 'xls':
    case 'xlsx': return 'excel';
    case 'ppt':
    case 'pptx': return 'powerpoint';
    default: return 'generic';
  }
}
