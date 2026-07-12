#!/usr/bin/env bun
import './lib/global-env.js';
import { captureCliException, flushGlitchTip, initGlitchTip } from './lib/glitchtip.js';
import { createM365Program } from './lib/m365-program.js';

const program = createM365Program();

(async () => {
  await initGlitchTip();
  try {
    await program.parseAsync(process.argv);
  } catch (err) {
    // Print a clean message (not the raw Error/stack) so an action handler that throws instead of
    // handling its own error surfaces something to the user rather than exiting silently.
    console.error(`Error: ${err instanceof Error ? err.message : String(err)}`);
    captureCliException(err);
    await flushGlitchTip(3000);
    process.exit(1);
  }
})().catch(async (err) => {
  console.error(`Error: ${err instanceof Error ? err.message : String(err)}`);
  captureCliException(err);
  await flushGlitchTip(3000);
  process.exit(1);
});
