# macOS shipment qualification

## Requirement and current boundary

The 2026-10-20 user request targets working Windows and macOS installers for
Wordplai. The user then confirmed they have no Mac and asked for this handoff
to a future agent on a Mac. Mac implementation and host qualification are deferred
to that agent. No Mac runtime, installer, minimum Word build, or installed Mac
result has been qualified. Windows remains the existing COM/WinForms add-in.
The active task checklist remains in `HANDOFF.md`; Windows publication must not
advertise Mac availability before the independent Mac gates pass.

## Start here on the Mac

Reuse this repository in the Mac's established GitHub checkout directory. Read
`AGENTS.md`, the `for-ai/README.md` router, `PROJECT.md`, `HANDOFF.md`,
`CROSS-PLATFORM-COMPATIBILITY.md` and the current verification ledger. Fetch and
inspect current Git state before resuming; this document does not pin a future
checkout to today's Windows candidate. Preserve unrelated work and documents.

Use disposable synthetic documents and retain ignored evidence with exact source,
package and installed payload hashes, commands, exit codes and actual host
observations. Record macOS, CPU architecture, Word version/build, supported API
sets and the intended distribution route. Establish a supported minimum from
observed capability, rather than guessing from documentation. Do not copy the
Windows private profile, DPAPI blobs, tunnel credentials or capability URLs.

Product owners are `modules/word/` and `modules/review/selective-editor-contract.json`.
Read existing native implementations and fixtures before choosing a Mac adapter.
Do not create a second permission model, an installer-only shell, or migrate the
working Windows stack. The API matrix below is the feasibility checkpoint: retain
actual failures and identify any missing safety capability before implementation.

## Candidate host integration

Microsoft documents the [Office Add-ins platform and Word JavaScript API](https://learn.microsoft.com/en-us/office/dev/add-ins/word/word-add-ins-programming-overview)
for Word on both Windows and Mac. This is a candidate Mac host boundary, not
evidence of Wordplai feature parity. Inspect the actual Mac Word build and runtime
requirement-set support before choosing APIs or claiming a minimum version.

The [desktop-only requirement-set guidance](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-desktop-1-4-requirement-set?view=word-js-preview)
requires runtime support checks and describes manifest limitations. Do not assume
that an API shown in the preview reference is available in the intended host.

| Existing owner / contract | Mac capability that must be proved |
| --- | --- |
| `WordTextMap`, `WordSnapshot` | Complete story/context reads and exact UTF-16 mapping, including fields, Unicode, tables and comments; no quote-search retargeting. |
| `WordMarkers`, `WordHighlighting` | Stable owned native anchors, exact RGB paint, rule persistence, overlap/protection and safe stale-marker reconciliation. |
| `WordNativeReview`, `WordProtectedContent` | Existing comment roots/replies and native units survive edits with their identities and formatting. `Range.insertComment` exists in WordApi 1.4; availability does not prove anchor preservation. |
| `WordEditor`, `WordNativeOutcome` | Precommit conflict checks, atomic refusal/rollback, truthful uncertain outcomes and native Undo semantics. A `context.sync()` batch must not be assumed to be an atomic transaction or custom Undo record. |
| `WordDocumentSelection`, `WordClientContexts` | Independent per-conversation document targets, duplicate-title isolation, close/reopen rules, and no active-window substitution. A document-scoped task pane alone does not establish application-wide document discovery. |
| `WordAutoSave`, `WordReviewStore` | Actual save/reopen, safe deferral, durable receipts, expiry/revocation and OS-appropriate protected storage. Windows DPAPI is not a Mac storage implementation. |
| `ConnectionService`, tunnel supervision | Real Mac-compatible connection binary, process ownership/cleanup and unchanged security boundaries. The bundled Windows AMD64 zrok executable is not a Mac runtime. |
| Native ribbon/dialogs | Mac-native or supported Office add-in controls, keyboard/focus/accessibility, resizing, full text access and equivalent user outcomes. |
| `WordAutomation`, `modules/word/cli/` | Equivalent local CLI outcomes without desktop input: explicit document identities, JSON file/stdin requests, state/context, local rule/marker/review commands and connection lifecycle. Exercise AI writes through the same four MCP tools with separate session credentials, complete-read witnesses and original request identities. Windows COM/PID binding and the `.exe` are not Mac implementations. |

## Smallest qualification sequence

1. Obtain a Mac with the intended desktop Word build and record its host/API
   capabilities on disposable synthetic documents. This requires the future
   agent's Mac and desktop Word installation; this Windows session has neither.
2. Prove reading, exact anchor/paint persistence, comments, one authorized edit,
   refusal, interrupted commit, Undo and save/reopen before expanding the adapter.
   Preserve the shared permission and uncertainty contracts; a missing safety
   capability blocks shipment rather than becoming a silent fallback.
3. Prove two distinct documents and independent client conversations through the
   actual transport. Preserve one authority per native document and existing
   request identities across reconnects.
4. Package the supported Mac integration and any native companion. Use Rust/Tauri
   guidance if a native companion is needed; read its architecture, security and
   release routes before implementing it. Do not migrate Windows as a prerequisite.
5. Run installation, update, removal and actual Word acceptance on the exact Mac
   package. Record CPU architecture, macOS/Word versions, package and installed
   payload hashes, signing/notarization state and real-client evidence.

Microsoft's [Mac sideload procedure](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac)
is a development validation route. Copying a manifest to `wef` is not proof of a
production installer or a qualified Wordplai implementation. The [Office add-in
hosting requirements](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/add-in-requirements)
also need a deliberate deployment choice; do not silently install trust roots,
weaken TLS, or expose local services to make the task pane load.

Keep Mac evidence independent from Windows and Writer evidence. Any unverified
capability remains unverified; a Mac-looking installer must never conceal a
missing or unsafe host implementation.

## Required handback from the Mac agent

- A supported host/version/architecture matrix, capability results and explicit
  unsupported cases. Any unresolved atomicity, protection, identity or uncertainty
  gap blocks a support claim.
- A reproducible build and actual installer with licenses, declared dependencies,
  signing/notarization status, checksum and verified installed payload identity.
  Prove first install, upgrade preserving local rules/receipts, uninstall and
  reinstall, including safe behavior while Word is open.
- Installed Word results for the scenario matrix in `HANDOFF.md`: local rule
  ownership, exact reads, authorized direct wording/rationale/green protection,
  protected native units, refused mutations, rollback/Undo, replay/uncertainty,
  two documents/conversations, revocation, real autosave and save/reopen.
- Real Chrome-extension WebGPT results on the Mac, including cold/reloaded picker,
  explicit choice, independent targets, safe failure recovery and measured first
  visible edit. Component tests and manifest sideloading cannot replace these.
- Installed CLI checks without mouse/keyboard input, including exact UTF-16
  targets/quotes, stale and closed-target refusal, independent sessions, native
  selection preservation and timer saving of local-only inactive changes before
  any MCP read. Local mutation must not grant remote authority. Reuse the Windows
  command/reference and scenario contracts while adapting their host binding;
  retain the separate physical and visual acceptance scopes.
- Updated existing verification/handoff owners, coherent commits, exact remote CI
  and download/deployment checks if publication is authorized. Publish sanitized
  evidence only; keep credentials, document identities and raw captures private.
