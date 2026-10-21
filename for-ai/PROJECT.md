# Project contract

## Purpose

A standalone Microsoft Word add-in that connects AI writing assistants to locally marked passages with user-owned color rules.

## Primary goal

The 2026-10-21 Windows milestone is full installer readiness with a working MCP
runtime on the user's PC and a locally saved Setup preference for the author
of future manuscript edit comments. Resume its acceptance in `HANDOFF.md`;
packaging or local installation alone does not close the full release goal.

Finish the full standalone acceptance matrix in `HANDOFF.md` and maintain
feature parity with LibrAI, correcting defects and retaining actual evidence on
each host. The public repository, installer and two-page site already exist.
Publication alone does not complete either goal.

The 2026-10-20 shipment goal additionally requires as much real installed/native
and Chrome-extension-driven WebGPT qualification as available, fixes for observed
defects, and working Windows and macOS installers. The user subsequently deferred
Mac implementation and qualification to a future agent with Mac/Word access;
this session must provide the actionable handoff in `MACOS-COMPATIBILITY.md`.
Retain the Windows native implementation. Mac support still requires a compatible
host integration and independently observed installed Word-on-Mac results;
packaging the Windows COM binary does not satisfy that requirement. The active
shipment checklist is in `HANDOFF.md`.

Completion requires passing component and hosted CI, native dialog geometry,
installed Word behavior, independent real WebGPT document sessions, cold/reloaded
picker rendering, physical navigation, measured edit latency and saved/reopened
state. Verify final packaging, downloads, Pages and remote synchronization.
Failed, blocked or unrun required cases remain incomplete; never substitute a
simulation or a component result for the named real surface.

## Cross-platform goal

Maintain feature parity with [LibrAI](https://github.com/GeorgeFejer91/LibrAI): every user-facing function in either sibling must have an equivalent user outcome on both supported hosts. LibreOffice may require a different native implementation protocol. The shared behavior, user authority, safeguards, and verification bar are defined in [`CROSS-PLATFORM-COMPATIBILITY.md`](./CROSS-PLATFORM-COMPATIBILITY.md); platform-specific implementation details remain with their own source and tests.

## Non-goals

- No speculative framework, service, abstraction, compatibility layer, or
  deployment system.
- No second implementation tree or duplicate source of truth.
- No capability claim without matching evidence.

Add project-specific non-goals only after they prevent a plausible wrong turn.

## Product/control-plane boundary

- Product source, assets, tests and packaging: `modules/word/`.
- Local Windows CLI and command reference: `modules/word/cli/`.
- Local selective-editor contract: `modules/review/selective-editor-contract.json`.
- Branding: `assets/wraiter/` (legacy asset names retained).
- Description and stable download pages: `site/`; deployment: `.github/workflows/`.
- Agent orchestration and durable project memory: `for-ai/`.
- Local generated diagnostics and scratch evidence: `.for-ai-local/` (ignored).

## Architecture and ownership

Keep the existing C# 5/.NET Framework 4.8 in-process Word COM stack and native
WinForms UI. Do not migrate to Tauri or require UnderTree. Retain legacy COM,
assembly and profile identities for compatibility; they are not an external
application dependency. The bundled zrok child has no inherited account state.

The Windows CLI uses the same .NET Framework/COM authority and is bundled with
the add-in. It provides explicit document-ID local automation and an MCP client;
it adds no remote tools or direct AI-write bypass. Local range marking shares
the native marking implementation. CLI tests require no mouse/keyboard input;
they do not replace physical-gesture, pixel or real WebGPT acceptance evidence.
Local `diagnostics` exposes numeric in-memory HTTP/RPC/resource counters for
testing resource delivery. It records no payloads or identities, requires no
document choice, and adds no remote tool. A completed response write does not
prove client receipt or widget bridge initialization.
Successful local CLI mutations enroll their explicitly chosen inactive document
in the existing native autosave path without sharing MCP authority. Reads and
refused mutations do not enroll it; ordinary named/writable, busy and input/save
guards continue to govern persistence.

Node.js is build/test tooling for the embedded picker and static website, not a
desktop runtime dependency. Pin Pretext and the bundler in the repository lock.
The public site has Description and Download navigation. `/wordplai/download/`
is stable across releases; installer assets are hosted in this repository.

Only locally assigned prompt passages permit wording edits. Direct wording
becomes green/protected with a concise native rationale. Neutral AI marks grant
no rule, and only the user assigns or changes color functions. Existing tracked
receipts and restrictive Discard operations retain local review. Preserve
protection, complete reads, selected-document isolation, uncertainty, replay,
input guards, native transaction rollback and automatic-save behavior.

## Current verified state

- Fresh Git repository initialized on 2026-10-18.
- AI control plane created and mechanically checked.
- Product extraction preserves all 40 native runtime source files and the
  qualified embedded picker. Independent build tooling, local shared contract
  and notices are included; private documents and connection state are excluded.
- The source delivery has outstanding foreground-editing, real WebGPT timing,
  host rendering and physical gesture acceptance gaps. Route these to the
  explicit handoff; publication does not close them.

Git and runnable checks are the authority for branch, revision, and behavior.
Do not turn this section into a second status ledger.
