export function checkReadOnly(cmdOrOptions?: any) {
  let isReadOnly = process.env.READ_ONLY_MODE === 'true';

  if (cmdOrOptions) {
    // If it's a Commander Command instance
    if (typeof cmdOrOptions.optsWithGlobals === 'function') {
      if (cmdOrOptions.optsWithGlobals().readOnly) {
        isReadOnly = true;
      }
    }
    // If it's just an options object
    else if (cmdOrOptions.readOnly) {
      isReadOnly = true;
    }
  }

  if (isReadOnly) {
    console.error('Error: Command blocked. The CLI is running in read-only mode.');
    process.exit(1);
  }
}
