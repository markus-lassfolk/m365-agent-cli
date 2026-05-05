import { afterAll, describe, expect, test } from 'bun:test';
import {
  prepareProgramForHelpVerify,
  teardownProgramAfterHelpVerify,
  verifyAllCliHelpAndDocExamples
} from '../lib/cli-help-example-verify.js';
import { createM365Program } from '../lib/m365-program.js';

describe('CLI help and CLI_REFERENCE examples', () => {
  const program = createM365Program();

  afterAll(() => {
    teardownProgramAfterHelpVerify(program);
  });

  test('every m365-agent-cli snippet resolves to a command and uses real flags', () => {
    prepareProgramForHelpVerify(program);
    const { errors } = verifyAllCliHelpAndDocExamples(program);
    if (errors.length > 0) {
      console.error(errors.join('\n\n'));
    }
    expect(errors).toEqual([]);
  });
});
