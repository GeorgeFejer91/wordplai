# Current qualification scope

The user asked to continue the standalone acceptance checks and fix concrete
gaps. Preserve the published preview and all unrelated applications/documents.

This pass owns the existing native document-choice fixture, the verification
workflow's artifact-action runtime pin, and the sanitized verification/handoff
updates. The 2026-10-19 focused Word run passed 50 assertions with zero drift
across 46 captured source/harness inputs. It used its separately spawned Word
process and disposable documents; the installed add-in and real ChatGPT
connector were not involved. No product runtime defect surfaced.

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
