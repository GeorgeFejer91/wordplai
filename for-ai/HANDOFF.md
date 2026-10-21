# Remaining Word and real-client qualification

## Active installer completion goal — 2026-10-21

The user requests a full Windows installer release and a working MCP runtime on
this PC, plus a Setup setting for the author of manuscript edit comments.
Keep this goal open until the release gates pass. Current work starts from
`3b19cb7` on main/origin/main; preserve the unrelated broken-name Git ref.

- [x] Persist a local comment author name (default wordplai), used only for new
  rationale comments; validate input, failed saves, legacy profiles and restart.
- [x] Verify Setup save/draft behavior, native control keyboard order and geometry at
  supported minimum and enlarged text, without internal panel scrolling.
- [x] Verify actual native comment attribution and DOCX reopen with an effective
  hardcoded-author fault control; preserve original manuscripts/global Word name.
- [x] Pass final component gates and exact installer/package checks; install the
  distinct candidate and verify the setting through installed Word/MCP.
- [x] Leave the installed MCP endpoint running and verify its four-tool protocol
  and a disposable document edit, preserving the original Word session.
- [ ] Resolve or explicitly obtain the required external evidence for remaining
  reliable client rendering, physical keyboard, visible latency and clean-machine
  release gates. Mac implementation remains deferred by the earlier user decision.

Writer parity for the new author preference is unimplemented/unverified; the
Windows feature does not establish cross-platform parity. The sibling checkout
is unavailable at the standard local path; do not claim its contract was updated.

Preview.9 is installed locally, with a distinct unsigned candidate under private
`artifacts/word/comment-author/candidate/`. Its 72 inputs/eight payloads, both
registrations and exact 150-file profile preservation pass. The seven original
open files are restored unchanged; Word PID 39856 owns the running public MCP
connection. Recheck it before use; copy the current address from Connections.
The default comment name is restored to wordplai after the installed test.
The installed background run passes 45 checks, including independent actual MCP
edits and exact configured authors in both native comments and timer-saved DOCX.
The first foreground run is separately PARTIAL with 18 checks and no positive
write. Public initialization/four tools/exact 58,378-byte picker delivery pass;
these do not qualify a host renderer. Full details and fault controls are in
`VERIFICATION.md`; retain initial restore/probe/layout failures privately.

No enabled Chrome/native UI surface is available in this session (only empty
in-app/MCP browser surfaces), and Windows Sandbox is absent. Clean Windows/Word
test-machine availability was requested. Next: use that environment for the
clean-install/uninstall gate, and connect the supported real Chrome/native UI
test surface for the outstanding host rendering, keyboard and timing checks.
Do not call the candidate a completed stable release or replace public preview.8
without the applicable gates. This request prepares the candidate; no new public
installer, site update or remote CI result is claimed.

## Exact RGB recolor follow-up — 2026-10-20

Exact green-to-red recoloring preserves text, ranges and existing rules. Neutral
red requires no editing prompt; hue never assigns a function. Policy 201,
transport 2,116, isolated native 217 and twenty-one component gates pass with
scoped evidence in `VERIFICATION.md` and private `.for-ai-local/recolor-20261020/`.
These source results do not qualify installed/live clients or Writer parity.
The independently committed CLI work `0aeb646` is preserved; later installed
and actual WebGPT outcomes are recorded below. Overall acceptance remains open.

This is a sanitized continuation contract for the standalone project, prepared
2026-10-18. It describes unfinished acceptance work, not a transcript or a claim
that publication completes that work. Record new results in
[`VERIFICATION.md`](VERIFICATION.md); keep private raw evidence out of Git.

The source-linked test strategy and risk-ranked stress shortlist are in
[`WORD-STRESS-TEST-RESEARCH.md`](WORD-STRESS-TEST-RESEARCH.md). Its three
complementary layers are state/transaction tests with retained fault controls,
isolated interactive Word fixtures, and installed/live-client observation with
accessibility and human gesture checks. They qualify distinct boundaries; do
not promote evidence from one layer to another.

## What must stay true

- Word wording has one mode: eligible locally marked `prompt` passages edit
  directly. There is no Suggestions/direct UI toggle. A successful direct edit
  adds one concise native rationale comment and becomes green Preserve text.
  Legacy prompt rules behave this way without rewriting their saved definitions.
- Existing pending/tracked receipts never become new direct edits. Discard
  remains a complete-removal proposal requiring the user's native Accept/Reject.
  Unresolved native revisions cannot be silently accepted by a new request.
- A fresh arbitrary RGB creates a neutral mark. An existing color identity may
  extend its unchanged local rule, including Protect, Reject, Discard or Prompt.
  Explicit recoloring replaces only the named exact current marker bindings.
  Only the local user defines or changes functions and prompts. Color operations
  do not rewrite prose or alter rule definitions.
- Complete current context, exact UTF-16 quotes/ranges, selected-document
  isolation, current protection, native identity, rollback and replay guards
  remain mandatory. A prompt or comment cannot grant additional authority.
- The add-in owns no remote Save or Accept operation. Ordinary native saves and
  the local autosave coordinator persist eligible documents. Never substitute
  a forced save, another document, a fresh request identity or relaxed protection
  for an unresolved result.

## Evidence boundary at extraction

The predecessor UnderTree tree has retained passing evidence for a 20-gate
component run, a 209-assertion isolated native workflow and a focused
27-assertion direct-comment workflow. These are overlapping scopes, not counts
to add together. Earlier failures remain part of that evidence history.

All 40 extracted C# runtime files were independently compared with the local
extraction hash manifest and matched. That establishes source identity only.
The inherited results do **not** establish that the standalone dependency lock,
fresh build, installer, installed runtime, Pages deployment or live connector
passed. Do not publish private predecessor manifests or relabel them as new
standalone results. Require fresh source/payload bindings for new checks.

Existing drivers are under `modules/word/tests/`; `modules/word/test.ps1` is the
component entry point. Run from this checkout after its extraction dependencies
are complete. Read each driver's scope before using it: an isolated source
fixture, a component control, an installed add-in and a real AI client are
different observations. A simulated gesture or bridge is not a physical gesture
or a successful host render.

## Known unresolved observations

**Retained CLI autosave regression:** preview.4 left inactive local rule/mark
changes dirty until an MCP read. The source regression fails before the fix and
passes 38 assertions afterward. Successful local mutations now enroll in native
autosave without sharing MCP authority; reads/refusals remain excluded. Preview.5
and the current installed preview.8 independently qualify persistence. Never
transfer those results to the retained older candidate.

**Foreground editing:** the current direct path refuses when the relevant Word
window owns foreground focus: "Finish typing and move focus from the Word
document before a direct AI edit." This is an actual current guard, not a
verified resolution of editing while typing. Test its precise pre-mutation
behavior and any proposed safe improvement; do not remove the guard merely to
make a demo pass.

**Original stale highlight report:** visible editing permission and a stale MCP
marker disagreed. The original failed edit's exact tool-error payload was not
captured, so stale metadata was a suspected cause, not a proven rejection cause.
Owned-bookmark remapping has since received focused qualification. A controlled
native-MCP fixture with an exact surviving bookmark and visible direct-prompt
color, but stale cached marker metadata, returned `context_changed` with
`request_applied=false` and a complete current context. It made no native or
saved-file changes; retrying the same request identity with that context applied
the exact wording, rationale and green protection. This verifies recovery for
that controlled state, not that it matches the original user trigger. Reproduce
the original user-visible discrepancy on a disposable fixture and retain the
entire actual error before attributing it to a particular subsystem.

**Host rendering:** real-client picker cards have failed during first load or
reload, including cases where the server's resource path returned successfully.
That does not prove client receipt, bridge initialization, rendering or cause.
Plain host/answer-load failures were also observed without a proven common root
cause. Keep these surfaces separate. Local widget tests and successful HTTP
responses do not close the live-host gap.

**Installed package identity:** each candidate requires its own payload, profile,
registration and installed-runtime evidence. Earlier preview.3/preview.4 outcomes
and failures remain in `VERIFICATION.md`; current installed preview.8 is below.
Do not transfer results between candidates or hosts.

**Retained hosted native UI failure:** run `39126772674` fails containment on a
1024×720 desktop despite a local pass. Commit `9ce27dc` adds readable no-fit
behavior, preserves drafts and restores editing when space returns. The old-code
regression fails; twenty local/hosted gates pass (`39175387165`). Full dimensions,
counts and failures remain in `VERIFICATION.md`. This native fix does not qualify
ChatGPT rendering, every installed environment or physical keyboard access.

## Prioritized scenario matrix

Use synthetic documents and independent request identities. Unless a row says
otherwise, verify both the intended effect and the absence of every forbidden
effect. Reuse existing fixtures before adding a new harness.

| Priority / scenario | Expected observable outcome | Evidence and timing points |
| --- | --- | --- |
| P0 — standalone build and package | Clean checkout builds with its own pinned tooling and local contract; installer includes the intended tunnel binary and licenses. No UnderTree checkout, private profile or account state is required. | Exact source, dependency-lock, picker, DLL and installer hashes; build exit codes; verify installed DLL against the package before any installed claim. |
| P0 — native dialog containment across Windows environments | Custom prompt and every other visible native control fit their allocated cells without clipping, internal scrolling or overlap. When the available desktop cannot fit readable editing controls, a readable Close state preserves the draft, disables hidden Save and restores editing when space returns. | Source regression, local and hosted containment now pass; preserve the original hosted failures and require installed/physical evidence separately. Do not skip the native UI gate or count browser picker tests as evidence. |
| P0 — direct wording and foreground | One authorized paragraph changes directly, with no native revisions, one rationale and immediate green protection. Word typing/focus refusal makes no edit and returns control promptly. Preserve and unmarked neighbors remain exact. | Request sent, dispatch admitted, first native text change, first visible rationale/green, response and later autosave. Exercise foreground Word, another foreground app and a retained inactive target without focus/input injection. |
| P0 — visible permission versus stale metadata | Unambiguous surviving owned anchors reconcile to the current visible permission; missing, ambiguous, revoked or genuinely stale restrictive authority refuses precisely. No remove/reapply workaround is counted as automatic repair. | Full read and failed/successful tool result, current native range/text/paint, marker status and revision. Distinguish each actual rejection reason; preserve protected text exactly. |
| P0 — requested colors and rule ownership | Fresh RGB is neutral; uniquely matching RGB or exact existing identity extends the unchanged rule for every supported action. Ambiguous same-RGB identities require an exact choice. Local assignment and later reuse behave consistently. | Exact requested/native RGB, new marker identity/ranges and before/after rules, text and formatting. Exercise citations and ordinary prose; no citation-specific shortcut. |
| P0 — recoloring and atomic batches | Explicit green-to-another-color recoloring replaces only selected identities. Identical, partial and nested selected overlaps are supported where authorized; unselected conflicting protection refuses atomically. Test 1, 9, 64 and 65 ranges, wrong final quote, duplicate IDs and unsupported final target. | Whole native/policy state equality on refusal; receipt replaced IDs, exact native paint and Undo/Redo on success. Record batch first-write and completion times without increasing deadlines to conceal a failure. |
| P0 — two independent conversations | Two real chats choose different open documents and can progress independently through the one endpoint. Fresh transports for one valid conversation retain its target and context; unrelated/no-metadata clients remain isolated. | Map synthetic target identities to native documents privately; prove exact wording/comments/receipts in each target and unchanged other documents. Interleave requests, response failures and window activation. |
| P0 — document choice and lifecycle | A one-document first read is pickerless. Multiple unnamed choices require one explicit selection; duplicate titles stay distinct by returned identity. A selected target stays selected after other windows open/activate. Close refuses without retargeting; explicit read may rebind a verified reopened instance. | Discovery contains names/identities without paths; no text or write before required choice. Test an inactive target, a late duplicate/copy and unrelated corrupt metadata. Stop clears retained authority. |
| P0 — context, retry and revocation | Changed source, feedback, rules or complete-read witness refuses the old request without mutation. Exact replay returns a receipt without duplicate edits/comments or cached successor context. Revocation and expiry cancel the appropriate authority. | Preserve full structured errors. Test fresh-context recovery, closed targets, transport/logical scope isolation, FIFO conflicts and queued cancellation. For uncertain outcomes, inspect the original receipt/native result; never retry under a new identity. |
| P0 — protected units, Unicode and discussion | Protected native text, fields, links, formatting and bookmark identities survive authorized enclosing edits intact. Each mapped unit occurs once. Comment roots/replies retain content, author, date, resolved state and parentage; fresh anchors identify the revised paragraph. | Emoji, combining text, half-surrogate boundaries, continuous/separate protection and protected suffixes. Test a shared ending, no shared ending, narrow anchor, replay and one-step Undo. Volatile export IDs and COM aliases are diagnostic, not persistent identities. |
| P1 — native structures and complete reading | Ordinary tables remain readable context; before/after prose has correct paragraph/comment offsets. Table/field/object write restrictions refuse atomically rather than flattening content. | Retain saved/reopened fixtures and exact native text/structure. Include a batch whose final target is unsupported, not just a successful outside-table range. |
| P1 — real autosave and reopen | The actual coordinator saves eligible named writable chosen documents, including inactive targets, fairly. Reopen preserves colors, rules, bookmarks, discussions and receipts. Accepted legacy tracked wording becomes green; Reject restores the original authority. | Observe timer activity, native Saved state and saved-byte changes without using a read or remote Save to trigger success. New unnamed documents, busy Undo, read-only and save failures defer safely. |
| P1 — real client role and minimal invocation | In two fresh real web chats, only "use wordplai MCP" starts the existing saved writing prompts after full context. The assistant acts as editor, submits each ready paragraph immediately, avoids capability narration/redundant questions and uses returned current context for the next edit. | Record client/model, initial writing role/instructions, connector/runtime identity and whether a document choice was required. Measure request-to-first-visible-edit separately from whole-task completion and upstream model delay. |
| P1 — picker first load, reload and fallback | Actual host initializes the bridge and renders a usable document choice on cold load and reload. Explicit named selection works without changing targets. A failed picker produces a truthful brief names-based fallback. | Separate server resource preparation/write, client receipt, bridge messages, first render, click result and first successful chosen read. Record plain-host failures independently. A server 200 or empty iframe does not identify the cause. |
| P1 — physical native navigation | A real ribbon/text double-click opens the exact selected color rule; overlaps require explicit choice. Keyboard access works. Opening, changing drafts and Cancel do not save; explicit Save alone changes the rule. No mode toggle returns. | Human/native physical gesture evidence plus exact marker/rule before/after state. Existing synthetic navigation tests do not qualify this physical interaction. |
| P2 — bounded stress and recovery | Document/session/client/mark/text/XML limits refuse without eviction, cross-target writes or an unresponsive editor. A fresh valid request recovers after a known safe refusal. | Exercise limit, one below and one above; sustained interleaving, expiry, Stop/reconnect, long context, late native callbacks and save failures. Preserve measured timings and original errors; do not weaken predicates to obtain green. |

## Error handling and evidence

Retain the exact request privately and capture the full response envelope,
`isError`, error code/message/details, request identity, operation-started and
outcome-unknown fields. For native diagnostics retain stage, exception type,
HRESULT, commit status and rollback attempted/confirmed fields as returned.
Do not replace an actual COM failure with a guessed stale-marker explanation.
If a field is absent, record it as absent rather than inventing a value.

`operation_started=false` with a known canceled queued dispatch is different
from a running timeout or `outcome_unknown`. The latter can still be completing
or committed. Preserve the original identity and native state; a new-ID retry
would bypass the uncertainty contract. Report a blocker once, without retry
loops. Keep successful prefixes of failed suites scoped as prefixes.

Use fresh ignored evidence directories, source/harness/package hashes, exact
commands, exit codes, actual conditions and a bounded meaningful negative
control. Never alter production while a source-bound check is running. Run
owned isolated Word fixtures serially where necessary; bind only the spawned
PID, never fall back to normal Word. Do not close or modify unrelated documents.
Installed and live-client checks need their own coordinated preservation plan.

Public reports contain only synthetic fixture descriptions, normalized relative
source paths, hashes, timings, scope and sanitized results. Exclude user document
names/content/IDs, capability URLs or tokens, credentials, local profile paths,
machine-specific paths, chat identities/links, raw native diagnostics, live DOM
captures and screenshots. Ignored files are not safe to attach automatically.
Keep unsanitized evidence private; do not overwrite it with a redacted copy.

## Readiness and next action

### Active shipment goal — 2026-10-20

The user authorized sustained Windows qualification, real Chrome/WebGPT
scenarios, fixes and qualified shipment. They subsequently deferred Mac
implementation/testing to a future agent with Mac and Word access and requested
an actionable handoff. The user subsequently explicitly authorized opening and
saving Word as needed for qualification. Preserve original document content and
audit lifecycle changes. Overall acceptance remains open.

- [x] Recheck preview.2 source/payload identity, install it with Word closed and
  verify its installed payload; run the installed 69-check fixture and preservation
  audit. These results cover preview.2, not the newer candidate.
- [x] Run eleven fresh isolated native suites serially, with retained failures and
  fault controls. Additional native autosave and color-input checks ran; the
  color-input process reporting defect is fixed: fresh native PASS/exit 0 with
  both effective mutant failures retained; existing-evidence refusal still exits 1.
- [x] Correct obsolete Connections overview text after a demonstrated regression;
  all twenty fresh component gates pass, including native UI containment.
- [x] Build preview.3, bind all 67 inputs and six payloads, and verify its exact
  DLL's COM ABI and effective fault control. See `VERIFICATION.md` for hashes.
- [x] Install preview.3, verify exact installed identity, upgrade/uninstall/reinstall
  and profile persistence; reopen all six originals with unchanged content,
  comments, revisions, saved state and saved-file hashes.
- [x] Complete preview.3's installed 69-check runtime fixture with all six
  originals unchanged. Retain the initial three-check cleanup refusal separately.
- [x] Qualify preview.7 physical ribbon/text double-click and preset draft,
  Cancel and explicit Save on a new owned fixture. Each stage passes 21 checks;
  ten independent saved-DOCX checks and an effective missing-preset control pass.
- [x] Qualify preview.8's corrected installed overlap chooser: exactly two rows,
  correct second protective rule, Cancel/Close and full preservation pass.
- [ ] Complete custom typing/keyboard. Earlier
  owned C was corrupted when targeted draft typing reached its document body;
  no Save occurred and native control was stopped by Escape. C stays invalid.
  The new nonmutating routing check also selects document text, so no new typing
  is attempted. The duplicate chooser correction is independently qualified below.
- [x] Reconnect an equivalent ChatGPT test app and run the first named-document
  real WebGPT edit; verify actual native/saved results and original preservation.
- [x] Run two independent real conversations and interleave color operations;
  verify protected-edit refusal, valid post-refusal recovery and native
  save/reopen with independent targets and original preservation.
- [x] Exercise real-chat local revocation, fresh-read recovery using the
  client-reported original request identity, and exact replay. Correlate the
  displayed native identities with saved state and the committed private ledger.
- [ ] Qualify a usable cold/reloaded picker, capture full client wire/error
  envelopes for uncertainty-field proof, and measure first-visible wording.
- [x] Prepare the future Mac agent handoff in `MACOS-COMPATIBILITY.md`.
- [ ] Future Mac agent: implement, package and independently qualify Word-on-Mac
  support. No Mac support or installer is currently claimed.
- [x] Publish Windows preview.3 as an unsigned prerelease with explicit remaining
  limits; verify all three asset digests and an anonymous installer download.
- [x] Verify the updated stable download page and exact remote CI/deployment
  after publication. Overall acceptance and Mac qualification remain open.
- [x] Qualify the newly requested CLI on its exact preview.4 package: local
  document/rule/range commands, inactive-target isolation, native selection and
  foreground preservation, MCP session/read/edit/replay, and original preservation.
  Native 33 assertions, client 14 checks and installed 16 checks pass. Upgrade,
  uninstall/reinstall and six-original preservation pass. The candidate is
  installed locally; preview.3 remains the public release. See `VERIFICATION.md`.
- [x] Extend installed CLI qualification with two sessions, unnamed choice,
  cross-client witness refusal, revocation/recovery, timer persistence and closed
  targets: 36 checks pass, with zero unqualified foreground comparisons. Twelve
  saved-file checks and effective copied session/rationale controls pass; all six
  originals retain native state and file hashes. Preserve the earlier failures.
- [x] Build/install and independently qualify preview.5's local-only background
  autosave fix before any MCP read: source-native 38, components 21 gates,
  exact-package COM 536, CLI 14, installed CLI 42 and saved-DOCX 12 checks pass.
  Session-routing fault control, Word-open installer refusal, upgrade/uninstall/
  reinstall and profile persistence pass. All 71 inputs/eight payloads are bound.
  Preserve preview.4's timer failure and source red/green receipts.
- [x] Publish the qualified Windows preview.5 after exact-source hosted CI;
  all three asset digests/sizes and anonymous downloads match. Preview.3 is
  preserved. Local stable-page rendering passes 313 checks with actual release
  metadata; verify its subsequent exact-SHA Pages deployment.
- [x] Publish qualified Windows preview.8 with explicit limits; all three public
  asset bytes match. Actual-release site rendering passes 313 checks; exact hosted
  CI/Pages and anonymous page/source/metadata audits pass. Older assets remain.
- [x] Qualify installed 1/9/64-range batches, 65-range/final-quote/duplicate-ID
  refusals, maximum recoloring and exact replay. Eighteen installed checks and
  eighteen saved-DOCX/receipt checks pass, with an effective missing-marker control.
  Native refusal revokes sharing; fresh complete reading recovers it. Preserve
  the original overstrict sharing assertion and all setup failures.

Latest installed state: public preview.8, with seven documents open
(five originals and owned G/H) and all six known original saved files unchanged.
The latest process is in `artifacts/word/preview8-installed-batches-isolated/cleanup-preservation-result.json`;
the previous process exited externally. Normal startup restored the originals;
indexed native/file preservation passes. Both previous-share cleanup and the
owned local endpoint Stop settle; no public endpoint is currently running.
An equivalent four-tool reconnect is required before new WebGPT tests.
The sixth original was already closed at upgrade
preflight; its cause is unestablished. Preserve that observed open set rather
than reopening it to satisfy the earlier six-open baseline. Native shutdown
initially exceeded the observer timeout; the same process later exited normally
before installer work. Reopen and post-test audits pass. Source and package
hashes, lifecycle receipts and limitations are in `VERIFICATION.md`.

At the earlier preview.5 checkpoint, the equivalent `wordplai CLI qualification` ChatGPT app was connected with the
same four tools and Allow low-risk permission. The installed public MCP endpoint
initializes with HTTP 200 in 1,001ms. Fresh independent Instant chats edit owned
G and H; G's first edit leaves H unchanged. The unnamed H chat's embedded picker
stays blank cold and after reload, then reports it could not display the app;
the assistant response also fails to load. Explicitly naming H continues the
same conversation and completes its edit. Twenty-two saved-DOCX checks verify
exact independent wording, Unicode, protected endings, green paint, one exact
rationale/anchor and applied receipt per document, with no revisions. Original
preservation still passes. Client task times are nineteen seconds for G and
eleven seconds for H's fallback; first-visible-edit latency remains unmeasured.
Browser control then detaches during preparation of a G-only color request;
no Send was performed, and both saved files match their pre-attempt hashes.
That detachment was recovered in fresh targets under the same Chrome browser.
The continuation in `artifacts/word/webgpt-cli-interleave/` privately qualifies
four color operations across G/H: neutral exact red and blue, then G red to
exact green and back to the existing neutral red rule. Forty-four saved-file
assertions confirm exact wording, other bindings/rules, native paint and
independent saved-byte isolation. A protected-ending attempt receives one-time
host approval and the client reports `INVALID_ARGUMENT` for missing native
preservation mapping; both saved files remain byte exact. The raw tool envelope
was not captured. A different valid color request then succeeds.

Only the two saved owned documents are closed/reopened hidden through the same
Word process. Eleven lifecycle checks, sixteen semantic comparisons and two
actual in-memory RGB checks pass: identities, native text/comments, rules,
markers, applied receipts and Saved state persist. Five open/six saved originals
remain unchanged. Retain the initial PowerShell Close binding failure, which
stopped before either fixture closed, and the initial comment XML byte check:
only volatile export IDs changed; complete content/attributes and mapped body
anchors compare exactly. The missing-neutral-color observation control fails
its intended predicate.

The public picker resource initializes/lists/reads with HTTP 200 and returns
the exact packaged 58,378-byte HTML. This does not prove ChatGPT receipt or
bridge/render initialization. The offered Open app in tab fallback also reports
could-not-display. No picker cause or code fix is inferred from those facts.
At that checkpoint real-chat revocation/recovery was still open; the subsequent
continuation below records its observed scope. Usable picker, first-visible
timing, physical UI and Mac scopes remain open. Public sharing was observed running;
recheck it before continuation. Keep private capabilities/browser evidence out
of Git and never retry an uncertain write under a new identity.

Latest revocation continuation: private evidence is in
`artifacts/word/webgpt-cli-revocation/qualification-result.json`. The retained H
chat reads its chosen document completely after reopen. The exact installed
local CLI changes shared=true to false without changing rules, marks or pending
proposals. One requested old-witness color attempt is client-reported refused
with `INVALID_ARGUMENT`; H remains unshared and both saved files stay byte exact.
After a fresh complete read, the client reports the original requested identity
and a first application of the color operation. Sixteen independent checks
correlate its returned native identities with the sole new committed receipt,
exact paint, unchanged wording/protection/rationale and unchanged G bytes.
The exact replay is reported as replayed=true with the same returned identities;
both saved files, the encrypted owned ledger and native/CLI review state remain
byte/structurally exact. Five open/six saved originals remain unchanged.

Highlight outcomes require both stores: document XML deliberately retains Undo
staging with a reserved outcome and precommit snapshot; the CurrentUser-protected
local ledger holds the applied outcome and final native snapshot. The runtime
merges those replay receipts on load. Retain both failed audit assumptions and
the effective copied client/native-identity mismatch control. Only the owned H
ledger was read; raw browser wire requests/error envelopes remain uncaptured.
The displayed error fields and request identity are client reports, not a claim
of raw transport capture.

ChatGPT required one-time approvals and flagged current tool prose/test IDs as
suspicious. Default persistent permissions stayed unchanged; no tool wording or
safety annotation was changed to bypass those warnings.

The subsequent picker follow-up adds `outputSchema` to
`list_manuscripts`, covering its exact choice/retained-selection result and coded
refusal envelope. This follows current OpenAI plugin guidance; general MCP makes
output schemas optional, so their absence does not prove the blank picker's
cause. The missing-schema regression fails against prior source; the corrected
transport passes 2,117 assertions and fifteen JSON Schema positive/negative
cases. Both real installed listing variants match the declaration, including an
explicit owned-document complete read and retained choice. All twenty-one source
component gates pass. Source checkpoint `3eed2b6` is pushed and passes hosted
verification `39500996679`.

Historical preview.6 contains that declaration. Word-open refusal, upgrade,
eight payloads, both registrations and exact 140-file profile preservation pass;
seven documents reopen unchanged. Its installed CLI remains PARTIAL: seventeen
checks pass while foreground Word prevents positive writing/isolation. Later
candidate passes never replace this result. Its 71-input package and raw receipts
are private under `artifacts/word/picker-output-contract/preview6/`.

The first helper closes Word before asynchronous Stop settles; the encrypted
ownership barrier survives unchanged, and normal Stop after reopening completes
cleanup. Await stopped/busy=false/cleanupPending=false before normal shutdown.
Keep the initial array-nesting refusal and corrected preflight. Old endpoints
are revoked; use current preview.8 process/status receipts, never a historical URL.

The equivalent four-tool preview.6 app fails cold/reloaded rendering; explicitly
named H still reads, and seven-open/six-saved-original preservation passes.
A corrected non-browser probe receives exact catalogue/58,378-byte HTML, while
the initial browser-like probe receives warning HTML. Neither establishes client
receipt/bridge initialization or cause. Version `51098c3` passes verification
`39504215085`, Pages `39504215244` and exact anonymous/Git-byte deployment audit.
Retain the Windows-line-ending/tag-label audit failures and effective hash control.

### Preview.7 rendering and overlap discovery — 2026-10-21

Historical preview.7 qualified 72 inputs/eight payloads, COM 536, CLI fifteen,
installed isolation forty-two and installed diagnostics nine checks pass. Settled
Stop, Word-open refusal, upgrade and seven-document restore pass; 140 profile
files remain exact. Version checkpoint `476342c` is pushed and passes hosted
verification `39511507796` and Pages `39511507524`. Anonymous deployed page/release
bytes and twelve exact-Git input hashes match; preview.5 was then public. Package
identities and retained source/CLI failures are indexed in `VERIFICATION.md`;
private evidence is under `artifacts/word/resource-diagnostics/preview7/`.

The equivalent four-tool app's first cold render fails, but reload visibly opens
seven choices and actual H selection resumes the conversation. A separate fresh
warm-cache conversation also fails its first render, including response-load
failure, then reload/actual G selection succeeds. Subsequent unnamed complete
reads interleave H then G and report their own retained targets; both preserve
the overlapping Keep unchanged rule. Final native/saved audit passes. Resource
read results/completed writes remain two through the warm case and interleaving:
the resource was cached. RPC/resource/Origin errors remain zero; aggregate HTTP
refusals are separate and do not identify the failure's cause. Raw bridge/client
wire is uncaptured. Keep both initial failures and reload successes distinct.
Private receipts are in `artifacts/word/resource-diagnostics/preview7-warm/`.

Installed physical ribbon and marked-text double-click open the correct rule.
A mouse-selected preset draft, Cancel, and explicit preset Save each pass 21
native/CLI/preservation checks. Ten independent saved-DOCX checks verify the exact
saved preset plus original custom prompt, marks, wording and RGB; a copied missing-
preset control fails. Custom typing remains unrun: the helper does not expose the
owned dialog as a selectable window, and a nonmutating Ctrl+A routing check selects
the document body. Alt/F10 attempts do not establish working keyboard navigation.
Mouse actions on the observed dialog screenshot region do work. Do not repeat the
earlier C corruption or type without verified dialog targeting.

The controlled overlap case finds a real defect: four selectable rows for two
identities, with a two-record pager; selecting/opening the repeated third row does
nothing. A source regression checks native LB_GETCOUNT before palette refresh and
reproduces native=4/managed=2. The shared list now queues resize/font/handle layout
after WinForms operations. Focused source UI passes 1,433 checks/21 renders,
including native row equality and exact selection after resize/font changes. A
copy of the original list fails the final regression. All twenty-one full source
component gates pass; preview.7 retains the observed defect. Private source evidence is
in `artifacts/word/overlap-list-fix/`, with physical receipts/screenshots in
`artifacts/word/native-ui-preview7/`.

Only the new disposable UI fixture is closed after its normal timer save. Its
chooser interactions leave policy/wording exact; the original seven-document set
and all six known original saved files pass final preservation. Source checkpoint
`8044b36` passes exact hosted verification `39515532203` before preview.8 packaging.

### Current preview.8 installed automation — 2026-10-21

Preview.8 is installed locally. All 72 inputs/eight payloads match; package COM
536/CLI fifteen, installed CLI forty-two and diagnostics nine checks pass, with
positive MCP writing and zero unqualified foreground checks. Settled Stop,
Word-open refusal, upgrade and exact seven-document restore pass; 143 profile
files remain unchanged. Exact identities are in `VERIFICATION.md`; receipts are
private under `artifacts/word/overlap-list-fix/preview8/`. Its later publication
is recorded below; earlier preview.5 assets remain unchanged.

Physical mouse selection now shows exactly two overlap rows and opens the second
protective rule correctly. Cancel/Close leave the owned fixture unchanged; it
closes normally. Custom typing/keyboard remains unqualified for the routing
reason above. No HTML/bridge or remote permission/tool change is included.

The equivalent preview.8 four-tool app still fails cold and reloaded rendering.
Explicitly named fresh I and J real chats each apply one saved edit. I displays
its success; J's response fails to load despite a committed edit. Reload also
fails. One read-only report retrieves the existing successful result without a
new write. Each actual saved DOCX passes fifteen independent checks, including
exact displayed Unicode wording, protective mark/paint, original rules/ending,
one applied receipt/rationale and neighbor byte isolation. A copied missing-
protection control fails. Raw wire is uncaptured. Never resend an unknown write
under a new identity because the browser response failed to load.

The read-only report leaves saved bytes, native state and full CLI state exact.
Only owned I/J then close, with unchanged saved bytes. Final original seven-open/
six-known-saved-file preservation passes. Keep initial observer failures separate:
PowerShell 5's implicit UTF-8 decoding mismatches the same JSON values, and a
prior UI-fixture cleanup path substitution fails before closure. Corrected
observers pass. Current Word process/public status are in private preview.8
restore/status receipts; the old endpoint is revoked. Recheck before reuse.

A later real-chat closed-target case reports `document_closed` and
`request_applied=false` for the exact owned J, without fallback/retry. Fifteen
checks correlate all displayed error fields with an independent actual CLI MCP
refusal and preserve seven open/six known saved originals plus both closed owned
files. The CLI envelope has `isError=true`; the client report says that field is
absent and reports outer `INVALID_ARGUMENT`. This is a channel distinction,
not proof that the real MCP wire omitted `isError`. A changed-code control fails.

Only owned J reopens hidden with its original identity, rules/marks, saved bytes
and rationale. Explicit real-chat recovery reads it completely; a subsequent
client-reported offset-only read retains J while another Word document stays
active. Thirteen recovery/preservation checks and an effective changed-identity
control pass. Normal close restores seven documents with all originals unchanged.
The original J response now renders after reopening the conversation; retain its
prior load failures. Private receipts: `artifacts/word/webgpt-preview8-closed-target/`.
No runtime/package or public asset changes accompany these read-only tests.

Version/evidence checkpoint `ae583e7` is pushed and passes hosted verification
`39521101640` and Pages `39521101662`. Local site rendering passes 313 checks;
anonymous page/release bytes and twelve deployed source hashes match exact Git
blobs, including the effective changed-hash control. Private consolidated CLI
proof is `preview8/cli-qualification-result.json`; remote/deployment receipts
remain under the same ignored evidence root.

Preview.8 is published as an unsigned prerelease with explicit limits. The actual
release-site checkpoint `e23d1e8` passes verification `39522289127`, Pages
`39522289143` and 313 local rendered checks. Anonymous Description/Download/
release bytes match tested output; twelve deployed source hashes match exact Git
blobs, and the changed-hash control is effective. Site version, publication time,
URL, checksum and size match GitHub's actual published asset metadata.
Post-publication native/saved preservation again passes seven open documents and
six known saved originals, with zero changes. Receipts are in the release root.
Inspected Chrome controls expose console logs and page assets, without a network
trace capability; the response menu exposes no raw tool result. Two Statsig cache
warnings do not establish the rendering failure's cause. Next: obtain a supported
host resource/bridge trace before attributing the rendering failure to code;
continue native keyboard and timing acceptance when the named surface is available.
Reliable host rendering, keyboard/custom typing, raw client wire, first-visible
timing, clean-machine and future Mac implementation/qualification remain open.

## Continuing test constraints and public release

Latest visible-workflow continuation: fresh owned K receives one real Chrome
request and exact timer-saved wording. Fourteen independent DOCX checks,
native/CLI correlation and the effective missing-protection control pass. The
host answer fails to load after the native edit. Approval stalls and an operator
pause leave normal latency unqualified; retained frame bounds and stale UIA text
are scoped in `VERIFICATION.md`. Only owned K/focus fixture close; Chrome returns
to its prior minimized state. Seven open/six known saved originals remain exact.
Private packet: `artifacts/word/preview8-visible-timing/qualification-result.json`.
The prior evidence checkpoint `0bb9840` is on `main`/`origin/main` and passes
hosted verification `39527533299`; recheck the next evidence commit separately.
External name-clash renames removed canonical Git metadata/control directories.
Matching branch/index copies and 22 exact tracked blobs restore their canonical
paths without source-content changes; recheck Git before any publication.
No production/package/public asset changes. Next: obtain supported host wire/
bridge evidence for rendering and a clean approval-resolved timing run; preserve
the uncertain-write rule and the separate native keyboard/Mac handoffs.

Use `modules/word/cli/README.md` and `test-installed-cli.ps1 -WordProcessId PID`
for automation without desktop input. Use the current process from the latest
private cleanup/preservation receipt above; never attach an arbitrary active Word instance. Browser setup
and native foreground checks run serially. Any unqualified foreground comparison
keeps the result PARTIAL; do not bypass the guard to manufacture a writing pass.
Physical dialog/gesture, clean-machine and timing scopes remain open.

Keep the earlier C fixture invalid. Older E/F preparations and A/B test histories
do not replace the current G/H evidence; B's Codex connector result is not a
second WebGPT chat. Use fresh owned prompt passages where the existing wording
has already become protected. Keep old connector accounts intact. The current
equivalent reconnect is already authorized; preserve the same four tools and
permissions. Historical failures and handoff chronology remain in Git history,
with the durable evidence index in `VERIFICATION.md`; do not relabel them passed.

Latest public release is [preview.8](https://github.com/GeorgeFejer91/wordplai/releases/tag/v0.1.0-preview.8),
targeting `994f03473bade689e6ad5f58482ecb88e1e19faf`; that checkpoint passes hosted
verification `39521437016`. The 23,885,366-byte installer SHA-256 is
`e0f8d9456798d8507021a893e48421ace9c8ae7ea67cc3104b152ab75a8de423`.
All three public assets match anonymous downloads byte for byte; attached sums
and GitHub digests/sizes match. The remote tag resolves to the intended commit.
Preview.5's three asset digests/sizes remain unchanged. Raw release/download
receipts are private under `artifacts/word/preview8-release/`.

Actual release metadata and limit copy are deployed in the existing root README
and two-page site at `e23d1e8`, with CI/Pages and exact anonymous audits above.
Older deployment/package outcomes remain in `VERIFICATION.md` and Git history.
The native source correction `8044b36` and version checkpoint `ae583e7` passed
their own exact hosted gates before publication; package inputs are unchanged.
Retain older assets and disclose the remaining preview limits. The overall
shipment goal remains active for reliable cold picker, timing, physical,
clean-machine, remaining host/client and future Mac scopes.
