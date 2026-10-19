# Current qualification scope

The user asked to continue the standalone acceptance checks and fix concrete
gaps. Preserve the published preview and all unrelated applications/documents.

Initial state is clean at `1d21fbabce9a9b42de67d44d002e560026727375`. This
pass owns a focused isolated native-MCP regression for a stale cached prompt
marker whose current visible Word permission still authorizes direct wording.
Expected files are this session scope, `modules/word/tests/NativeReviewSmoke.cs`,
`modules/word/tests/test-native-review.ps1`, `for-ai/scripts/check-context.ps1`,
`for-ai/VERIFICATION.md`, and `for-ai/HANDOFF.md`.
The gate's remote check must retain Git's exit code before filtering its output;
PowerShell 7.6.5 leaves `$LASTEXITCODE` unset after the current pipeline.
Preserve the MCP request/response contract and assert the actual edit plus
native text, review state, local rules and saved-file effects. No product
runtime, installer, release, site, connector configuration or user document
may change. The earlier document-choice run passed 50 assertions with zero
drift; the prior native-review suite passed 326 isolated assertions. No known
pre-existing test failure is being carried into this pass.

Do not operate on the user's open Word documents. Preview.2 is still a local
candidate; install only after the user confirms Word is saved and closed.
Continue to distinguish source fixtures, installed Word, physical gestures
and real-client behavior.

Chrome side-session inspection reaches a fresh ChatGPT page, but the API does
not expose the Chrome toolbar avatar or profile selector. The page's ChatGPT
account icon showed orange initials, which does not confirm the yellow-figure
Chrome profile the user identified. Do not claim real-client qualification
until the requested profile and connector are available.

| Risk | Verification |
| --- | --- |
| A source fixture is mistaken for installed behavior | Keep isolated, installed, physical and real-client evidence separate in `for-ai/VERIFICATION.md`. |
| Wrong Chrome profile is used for live-client checks | Confirm the yellow-figure profile is active before using the real connector. |
| Artifact upload action runs on a deprecated runtime | Pin the official Node 24 action release and retain the existing artifact behavior; verify hosted upload and the full component workflow. |
