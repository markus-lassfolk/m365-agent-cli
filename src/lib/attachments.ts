import { lstat, realpath, stat } from 'node:fs/promises';
import { isAbsolute, normalize, resolve, sep } from 'node:path';

const MAX_ATTACHMENT_SIZE_BYTES = 25 * 1024 * 1024;

export interface ValidatedAttachmentPath {
  inputPath: string;
  absolutePath: string;
  fileName: string;
  size: number;
}

export class AttachmentPathError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'AttachmentPathError';
  }
}

function hasParentTraversal(inputPath: string): boolean {
  const normalizedInput = normalize(inputPath);
  return normalizedInput.split(/[\\/]+/).includes('..');
}

export async function validateAttachmentPath(
  inputPath: string,
  allowedBaseDir: string
): Promise<ValidatedAttachmentPath> {
  if (!inputPath || !inputPath.trim()) {
    throw new AttachmentPathError('Attachment path cannot be empty');
  }

  if (hasParentTraversal(inputPath)) {
    throw new AttachmentPathError(`Path traversal is not allowed in attachment path: ${inputPath}`);
  }

  if (inputPath.startsWith('~')) {
    throw new AttachmentPathError(`Home directory shortcuts (~) are not allowed for attachments: ${inputPath}`);
  }

  if (isAbsolute(inputPath)) {
    throw new AttachmentPathError(`Absolute paths are not allowed for attachments: ${inputPath}`);
  }

  const candidatePath = resolve(allowedBaseDir, inputPath);
  const realAllowedBase = await realpath(allowedBaseDir);

  let realCandidatePath: string;
  try {
    realCandidatePath = await realpath(candidatePath);
  } catch {
    throw new AttachmentPathError(`Attachment does not exist: ${inputPath}`);
  }

  if (realCandidatePath !== realAllowedBase && !realCandidatePath.startsWith(`${realAllowedBase}${sep}`)) {
    throw new AttachmentPathError(`Attachment path escapes the allowed directory: ${inputPath}`);
  }

  const symbolicLinkInfo = await lstat(candidatePath);
  if (symbolicLinkInfo.isSymbolicLink()) {
    throw new AttachmentPathError(`Symbolic links are not allowed for attachments: ${inputPath}`);
  }

  const fileInfo = await stat(realCandidatePath);
  if (!fileInfo.isFile()) {
    throw new AttachmentPathError(`Not a file: ${inputPath}`);
  }

  if (fileInfo.size > MAX_ATTACHMENT_SIZE_BYTES) {
    throw new AttachmentPathError(`File too large (>25MB): ${inputPath}`);
  }

  return {
    inputPath,
    absolutePath: realCandidatePath,
    fileName: realCandidatePath.split(/[\\/]/).pop() || inputPath,
    size: fileInfo.size
  };
}
