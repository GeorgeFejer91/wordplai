# Current implementation scope

The user requested completion of the entire standalone acceptance matrix.
Start from the existing public repository; preserve the preview's evidence and
all unrelated applications and documents.

The first change owns native rule-dialog geometry and its existing UI tests:
`modules/word/src/RuleDialog.cs`, `modules/word/tests/UiLayoutSmoke.cs`, the
component CI evidence step if needed, and their narrow For-AI status updates.
Shared native layout changes require evidence that the same defect lives there.

Initial state: clean `main`, commit `4ac5828`. Pre-existing failure: hosted
native UI allocates 115px to a 166px Custom prompt group; local UI previously
passed. Foreground editing, live host behavior, physical gestures and installer
qualification remain separate acceptance work.

Keep permission, coloring, direct-edit, protection, transaction and replay
contracts unchanged by the geometry fix. Do not edit UnderTree or other product
modules. Do not operate on unrelated open Word documents.

| Risk | Verification |
| --- | --- |
| Native controls overrun their cells or actions | Existing containment and readable-prompt assertions, added font/DPI/form metrics, retained renders, local and hosted runs. |
| Async sizing hides an unstable result | Observe actual native geometry convergence; retain the failure rather than adding a sleep or weakening bounds. |
| Sizing change alters rule authority | Existing native UI dispatch/Save/Cancel assertions and complete component suite. |
| A simulated check is mistaken for installed behavior | Keep isolated, installed, physical and real-client results separate under the canonical handoff. |
