export type RenameItemType = 'file' | 'folder';

export type RenameNameValidation =
  | { valid: true; value: string }
  | { valid: false; error: string };

const INVALID_NAME_CHARACTERS = /[\\/:*?"<>|#%{}~&]/;
const MAX_NAME_LENGTH = 128;

export function validateRenameName(value: string, currentName: string): RenameNameValidation {
  const trimmedValue = value.trim();

  if (!trimmedValue) {
    return { valid: false, error: 'Enter a name.' };
  }

  if (trimmedValue.length > MAX_NAME_LENGTH) {
    return { valid: false, error: 'Names must be 128 characters or fewer.' };
  }

  if (INVALID_NAME_CHARACTERS.test(trimmedValue)) {
    return { valid: false, error: 'Names cannot contain \\ / : * ? " < > | # % { } ~ &.' };
  }

  if (trimmedValue === '.' || trimmedValue === '..' || trimmedValue.endsWith('.')) {
    return { valid: false, error: 'Names cannot be . or .. or end with a period.' };
  }

  if (trimmedValue === currentName) {
    return { valid: false, error: 'Enter a different name.' };
  }

  return { valid: true, value: trimmedValue };
}

export function buildRenameEndpoint(siteUrl: string, itemType: RenameItemType, sourceServerRelativeUrl: string, newName: string): string {
  const sourcePath = sourceServerRelativeUrl.replace(/\/+$/, '');
  const parentPath = sourcePath.substring(0, sourcePath.lastIndexOf('/'));
  const destinationPath = `${parentPath}/${newName}`;
  const api = itemType === 'folder' ? 'GetFolderByServerRelativeUrl' : 'GetFileByServerRelativeUrl';
  const normalizedSiteUrl = siteUrl.replace(/\/$/, '');

  return `${normalizedSiteUrl}/_api/web/${api}('${escapeODataString(sourcePath)}')/MoveTo(newurl='${escapeODataString(destinationPath)}',flags=0)`;
}

function escapeODataString(value: string): string {
  return value.replace(/'/g, "''");
}
