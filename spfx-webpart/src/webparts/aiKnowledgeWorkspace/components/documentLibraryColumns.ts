export interface IDocumentLibraryDisplayColumn {
  internalName: string;
  title: string;
}

const nameColumn: IDocumentLibraryDisplayColumn = { internalName: 'FileLeafRef', title: 'Name' };
const requiredListItemFields: string[] = [
  'Id',
  'FileLeafRef',
  'FileRef',
  'FileDirRef',
  'File_x0020_Type',
  'Modified',
  'FSObjType',
  'Editor/Title'
];

export function parseDisplayFieldOrder(value: string | undefined): string[] {
  const seen: Record<string, boolean> = {};

  return (value || '').split(',')
    .map((field) => field.trim())
    .filter((field) => {
      if (!field || seen[field]) {
        return false;
      }
      seen[field] = true;
      return true;
    });
}

/** Converts the editor's ordered dropdown selections to one unique column list. */
export function getOrderedColumnSlots(slots: Array<string | undefined>): string[] {
  const seen: Record<string, boolean> = {};
  return slots
    .map((slot) => (slot || '').trim())
    .filter((slot) => {
      if (!slot || slot === 'FileLeafRef' || seen[slot]) return false;
      seen[slot] = true;
      return true;
    });
}

export const knownDocumentLibraryDisplayColumns: IDocumentLibraryDisplayColumn[] = [
  { internalName: 'File_x0020_Type', title: 'Type' },
  { internalName: 'Modified', title: 'Modified' },
  { internalName: 'Editor', title: 'Modified By' },
  { internalName: 'Created', title: 'Created' }
];

export function buildSelectedDisplayColumns(
  slots: Array<string | undefined>,
  discoveredFields: IDocumentLibraryDisplayColumn[]
): IDocumentLibraryDisplayColumn[] {
  const fieldsByInternalName: Record<string, IDocumentLibraryDisplayColumn> = {};
  knownDocumentLibraryDisplayColumns.concat(discoveredFields).forEach((field) => {
    fieldsByInternalName[field.internalName] = field;
  });

  return getOrderedColumnSlots(slots)
    .map((internalName) => fieldsByInternalName[internalName])
    .filter((field): field is IDocumentLibraryDisplayColumn => !!field);
}

export function buildDisplayColumns(
  availableFields: IDocumentLibraryDisplayColumn[],
  persistedOrder: string | undefined
): IDocumentLibraryDisplayColumn[] {
  const fieldsByInternalName: Record<string, IDocumentLibraryDisplayColumn> = {};
  availableFields.forEach((field) => {
    fieldsByInternalName[field.internalName] = field;
  });

  const selectedColumns = parseDisplayFieldOrder(persistedOrder)
    .filter((internalName) => internalName !== nameColumn.internalName)
    .map((internalName) => fieldsByInternalName[internalName])
    .filter((field): field is IDocumentLibraryDisplayColumn => !!field);

  return [nameColumn, ...selectedColumns];
}

export function buildListItemSelect(selectedInternalNames: string[]): string[] {
  const selectedFields = selectedInternalNames.filter((internalName) => internalName && internalName !== 'Editor' && requiredListItemFields.indexOf(internalName) === -1);
  return requiredListItemFields.concat(selectedFields.filter((internalName, index) => selectedFields.indexOf(internalName) === index));
}
