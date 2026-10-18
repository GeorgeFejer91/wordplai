# Project contract

## Purpose

A standalone Microsoft Word add-in that connects AI writing assistants to locally marked passages with user-owned color rules.

## Primary goal

Deliver an independently buildable Windows Word add-in, installer, public source
repository and two-page GitHub Pages site. Preserve known editing restrictions
and hand off unfinished real-client stress tests without claiming they passed.

## Non-goals

- No speculative framework, service, abstraction, compatibility layer, or
  deployment system.
- No second implementation tree or duplicate source of truth.
- No capability claim without matching evidence.

Add project-specific non-goals only after they prevent a plausible wrong turn.

## Product/control-plane boundary

- Product source, assets, tests and packaging: `modules/word/`.
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
