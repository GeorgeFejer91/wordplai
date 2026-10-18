# wordplai

A standalone Microsoft Word add-in for AI-assisted writing in your document.
Mark passages in Word, assign their color a writing prompt, and ask your connected
assistant to use Wordplai. The native add-in applies authorized wording directly,
adds a short rationale comment, and marks completed wording green/protected.

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
physical gestures and clean-machine installer/upgrade cases are not fully qualified.
The [remaining acceptance handoff](for-ai/HANDOFF.md) names the required tests.

## Build and test

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
