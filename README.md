# wordplai

A standalone Microsoft Word add-in for AI-assisted writing in your document.
Mark passages in Word, assign their color a writing prompt, and ask your connected
assistant to use Wordplai. The native add-in applies authorized wording directly,
adds a short rationale comment, and marks completed wording green/protected.

[LibreOffice counterpart: LibrAI](https://github.com/GeorgeFejer91/LibrAI) ·
[Description](https://georgefejer91.github.io/wordplai/) ·
[Download](https://georgefejer91.github.io/wordplai/download/) ·
[Preview releases](https://github.com/GeorgeFejer91/wordplai/releases)

## Your document, your rules

- Only you assign or change color functions and prompts in Word.
- An assistant can add neutral colors, extend an existing unchanged color rule,
  and recolor text when requested. Neutral color alone grants no wording access.
- Protected passages remain unchanged. Existing tracked edits and restrictive
  Discard operations retain local review.
- Each conversation keeps its chosen document. With several open documents,
  an unchosen conversation must choose one; one document needs no picker.
- Edits and automatic saving remain subject to native revision, input and
  document-protection checks. Unknown outcomes must be inspected, never blindly retried.

## Install the preview

Requires Windows 10/11 x64 (Intel/AMD), desktop Microsoft Word and .NET Framework 4.8. Close Word
before running the installer. Wordplai is installed per user and bundles its
tunnel executable; it does not require UnderTree, LibreOffice or Node.js to run.
Connection/account setup remains yours, and no connection credentials are shipped.

This is an unsigned early build. The existing direct-edit guard can refuse while
Word is the foreground application. Fresh WebGPT timing, host picker rendering,
physical dialog draft/Cancel/Save and clean-machine cases are not fully qualified.
Preview.3 fixes the color-dialog sizing failure on small desktops with enlarged
text using a readable no-fit state that preserves the draft. Its exact installer
passes 69 installed Word checks and upgrade/uninstall/reinstall on the qualification
host. Mac support and a Mac installer remain deferred to the
[future Mac agent handoff](for-ai/MACOS-COMPATIBILITY.md).
The [remaining acceptance handoff](for-ai/HANDOFF.md) names the required tests.

## Build and test

The [Windows CLI](modules/word/cli/README.md) automates local rules, range marking,
document-isolated state/context, connection lifecycle and MCP test requests using
JSON, without desktop input. It is included in source builds and preview.4
packaging; the earlier preview.3 installer does not contain it.

To exercise an installed build with disposable documents and no desktop input:

```powershell
./modules/word/tests/test-installed-cli.ps1 -CliPath "$env:LOCALAPPDATA/Programs/wrAIter Word/wordplai.exe" -WordProcessId WORD_PID -OutputDirectory artifacts/installed-cli-new
```

Use the PID of an already running Word instance with the add-in connected.
The driver creates and closes only disposable test documents. The positive AI
edit is skipped when that Word process holds foreground focus; the existing
input guard still applies. Fresh-process source tests are also documented in the
CLI reference.

Add `-IsolationCases` to test two independent MCP clients, cross-client context
refusal, revocation recovery, timer-saved files and closed targets. When using
the desktop during a run, `-AllowForegroundChanges` lets the behavior checks
continue and records changed-foreground comparisons as unqualified; any such
comparison makes the result `PARTIAL`. Native edit guards still apply.

Build tooling: Windows, PowerShell 7, .NET Framework C# compiler, Node.js 22 or
newer, and NSIS for the installer. See [the Word module](modules/word/README.md)
for exact standalone commands and pinned runtime provisioning. Dependency
versions are locked in `package-lock.json`.

The native source and original synthetic test suites live in `modules/word/`.
The selective-editor contract and branding are included locally. Legacy internal
assembly/COM/profile names are retained for upgrade compatibility; they do not
introduce a runtime dependency on the original application.

## Project context

[For-AI entry point](for-ai/README.md) · [Verification](for-ai/VERIFICATION.md) ·
[Remaining tests](for-ai/HANDOFF.md) · [Third-party notices](modules/word/THIRD-PARTY.md)

Project source retains its inherited `UNLICENSED` status. Third-party components
keep their stated licenses and notices; a new project license has not been assigned.
