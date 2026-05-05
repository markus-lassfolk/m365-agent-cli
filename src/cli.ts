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
    captureCliException(err);
    await flushGlitchTip(3000);
    process.exit(1);
  }
})().catch(async (err) => {
  captureCliException(err);
  await flushGlitchTip(3000);
  process.exit(1);
});
