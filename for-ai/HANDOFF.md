# Remaining Word and real-client qualification

## Exact RGB recolor follow-up — 2026-10-20

The existing recolor authority supports green `#00FF00` to red `#FF0000` with
unchanged text/ranges and existing rules. MCP guidance now explicitly permits
neutral red with an empty prompt and forbids inferring a function from hue or
asking for an editing prompt for a visual request. Existing red rules retain
their local meaning. Focused policy checks pass 201 assertions, including fresh
and existing neutral red and unchanged Reject red; transport checks pass 2,116.
The isolated native suite passes 217 recolor assertions, including exact paint,
other-marker preservation, replay and save/reopen. The final component driver
passes all 21 gates, including 1,419 native UI checks and 21 renders. Its overall
`PARTIAL` retains the installed/public-client evidence boundary. See the scoped
results and test-source line-ending receipt in `VERIFICATION.md`.
Receipts are private under `.for-ai-local/recolor-20261020/`; installed/live AI
clients and Writer parity are not qualified by these checks. Preserve the
pre-existing preview.4 CLI work, now independently committed as `0aeb646`.
This local change is based on `b79108a` on main; no installer or publication was
performed. Next acceptance remains installed/live-client qualification under
the existing checklist below; the overall goal remains open.

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

**Installed package identity:** the local preview.4 CLI candidate is installed
with all eight exact payloads verified. Upgrade, uninstall, registration removal
in both registry views and reinstall pass; all 112 existing profile files are
unchanged. Its installed CLI passes 16 checks and all six originals are preserved.
The published preview.3 separately passed its 69-check installed fixture after
normal Stop cleared cleanup and normal Chrome activation moved foreground focus.
Full hashes and scoped receipts are in `VERIFICATION.md`. Do not transfer results
between candidates or hosts.

**Fresh hosted native UI:** standalone workflow run `39126772674` passed policy,
transport and autosave checks but failed the native WinForms containment gate.
The Custom prompt GroupBox required 166 pixels in a 115-pixel-high TableLayout
cell at the tested size. The corresponding local 1,322-check run passed. A
fresh hosted capture identified a 1024×720 working area, 18.875pt text and a
screen-clamped minimum that could not contain the required controls. Commit
`9ce27dc` fixes this with a readable no-fit state, preserves the open draft and
restores editing when space returns. The regression fails against the old code;
all 20 component gates pass locally and in hosted run `39175387165`, including
1,419 local and 1,373 hosted native UI checks with 21 renders in each run.
The original failures remain retained. Published preview.1 does not contain
this fix; preview.2 has since been installed and tested as recorded below. Installed UI and
physical interaction qualification remain open. This is separate from the
ChatGPT picker and the earlier JSON property-order assertion correction.

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
- [ ] Complete physical dialog draft/Cancel/Save checks. A real double-click on
  owned C opens the correct installed dialog, but targeted typing actually
  replaced C's document body with the disposable draft string. The dialog image
  did not reflect that input. C is invalid for future tests; retain this failed
  interaction. No dialog Save was performed. Native control was stopped by Escape.
- [x] Reconnect an equivalent ChatGPT test app and run the first named-document
  real WebGPT edit; verify actual native/saved results and original preservation.
- [ ] Run a second independent conversation, cold/reloaded picker, independent
  targets, safe refusal/retry, revocation and first-visible-edit timing.
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

Chrome extension browser access was observed and one real connection-only
WebGPT request attempted discovery but failed internally. A reload temporarily
showed an empty assistant answer; a later native view showed the failed response.
Neither observation proves a picker failure or an add-in rejection cause. The
saved connector endpoint differs from the installed candidate's newly generated
endpoint. The user explicitly approved a replacement connector with the same four
tools and document permissions for synthetic testing. Reconnection completed as
`wordplai qualification`; its visible tool panel confirms four tools. Do not ask
for that same approval again or delete old connectors. The first named-document
Instant chat executed a saved prompt: exact saved text, native rationale, green
paint and stored protection pass fourteen saved-DOCX checks. The post-run audit
confirms all six original documents unchanged. The client showed nineteen seconds
for the whole task; first-visible-edit latency is not measured. DOM/screenshot
show the answer, although AX omitted its text. A second browser tab opened despite
a creation timeout, then control detached. B was subsequently qualified through
the public Codex MCP connector, not that second browser conversation: authorized
wording, protection/stale refusals, replay, neutral RGB, atomic quote refusal and
saved-state checks pass, with all six originals preserved. B's prompt wording is
now protected and its intro/ending have a neutral color. Use fresh owned passages
for future two-browser writing scenarios; do not count B's Codex result as WebGPT.

The public Word connection briefly reported running, then failed/revoked; HTTPS
returned a tunnel-hosted 404. Separate read-only synthetic probes through the
exact installed TunnelProcess/pinned zrok produced an informational browser
interstitial and, with a normal API client identity, 502/timeouts. A later verbose
probe reached the exact synthetic response at about 13 seconds and retained a
live child for another 15 seconds; exact owned-share cleanup passed. This proves
one synthetic transport success. Later actual public Word/MCP initialization and
the first real WebGPT edit passed. Reliability across restart/revocation remains
unqualified. No TLS/security bypass or expanded permissions were used; the earlier
failures' cause remains unconfirmed.

A subsequent installed connection-start request returned starting. The following
COM status query found no exposed add-in object; the existing Word process still
responds. A later probe confirms the add-in exposed/connected and public running;
the first real edit and post-run preservation audit then passed. Chrome control
remains intermittent. Do not infer shutdown from an observation timeout or repeat
a writing request under a new identity to compensate for lost browser control.

The previous bounded code scope was the Connections overview and its existing UI
assertion, plus correction of the color-input driver's inherited exit status.
Keep editing/protection/transaction/connection permissions unchanged. Preserve
the initial native UI resource-count failure independently; later passes do not
prove its cause. Do not restart a live test or alter its captured source inputs.

Hosted run `39427161161` on documentation checkpoint `b79688d` failed the recolor
state assertion because nested JSON object properties were ordered differently.
The captured values were confirmed identical with the existing structural
comparer and strict controls. The completed test-only repair shares it between
`PolicyTests.cs` and `PolicyRecolorTests.cs`, update their compile inputs in
`modules/word/test.ps1`, and retains values, fields and list order exactly. All
twenty component gates pass after the repair, and a scalar-losing copied control
fails the intended identity assertion. No runtime, permission, timeout, installer
or production serialization change. Hosted verification `39428499738` passed on
the correction's exact published SHA `7034ff0`; retain the earlier hosted failure
and partial local repair.

The user requested CLI functions so behavior testing can continue without taking
over the desktop. Use `modules/word/cli/README.md` and the installed CLI driver;
retain the separate physical/UI gaps. On the latest UI resume, Chrome's extension
session was absent, native browser URL verification stopped control, then a
Word-only double-click attempt was stopped by Escape. Fresh E/F fixtures were
prepared with explicit UTF-8 baseline serialization; C remains invalid.

Next actions: use the CLI and `test-installed-cli.ps1 -WordProcessId PID` for
repeatable native/MCP behavior tests without desktop input. The driver requires
an already running, connected Word instance; it creates and closes only its own
fixtures. A fresh-process installed launcher was dropped after retained startup
and third-party-template failures. Source-native isolated tests remain available.
Use `-IsolationCases` for the expanded installed cases. Foreground drift can be
diagnosed with explicit `-AllowForegroundChanges`; any unqualified comparison
keeps the result PARTIAL. The final 36-check run qualified all comparisons.
Chrome extension access is restored and the current qualification connector's
settings are readable. Its saved endpoint is obsolete after Stop. Reconnect the
same four tools under the existing user authorization, prepare fresh owned prompt
fixtures, and exercise two actual WebGPT chats. Browser session setup and native
foreground checks must run serially. No new WebGPT writing request has been sent.
Resume physical
UI work only after restored access. Prepare fresh owned
prompt documents; C was changed by the failed UI interaction and must not be
counted as an editing success. Qualify the dialog draft/Cancel/Save and interleave
two browser chats without cross-target effects, then exercise cold/reloaded
choice and safe failure/revocation. The second created tab was absent from the
latest tab inventory; the original A chat remained available. Preview.3's public
connection reached running after its installed fixture; public initialization
returned HTTP 200 in 629ms. Normal Stop completed cleanup, and the old URL then
returned HTTP 404. Sharing is stopped; future live-client work needs a new
connection and equivalent connector reconnection. Retain the foreground guard and
initial cleanup refusal. Its Word-open installer refusal
already passes with installed files unchanged. The qualified Windows preview.3
is published with the explicit preview limits under decision D-0002; site
metadata now identifies its exact asset, hash, size and actual publication time.

Source checkpoints: `9b768c1` corrects the overview/builds preview.3, `ba9ff8a`
corrects test reporting, and `40ba956` records qualification and the Mac handoff.
They, the shared-skill update, test-only comparer correction `7034ff0` and
installed evidence `6c455db` are pushed on `main`. Hosted verification
`39431955158` passed on `6c455db`, the preview.3 release target. The installer
remains bound to the unchanged 67 inputs and six payloads. Site/metadata checkpoint
`25a992f` passed verification `39433032039` and Pages `39433031883`. Anonymous
Description, Download and release metadata match tested output; all twelve
deployed manifest inputs match that published Git commit. Recheck remote SHA and
applicable CI after subsequent changes.

CLI checkpoint `0aeb646` is pushed; hosted component verification
`39470432974` passes. Its candidate matches all 71 committed inputs and the
installed driver. A test-only site readiness correction follows retained Pages
failure `39470433033`; local site verification passes 313 checks. Concurrent
recolor work in the current scope, module README, MCP guidance and three existing
test files is preserved outside the CLI commit. Recheck exact remote CI after
the follow-up. Preview.4 remains a local candidate.
The follow-up `b79108a` subsequently passes both hosted verification
`39471037155` and Pages `39471037096`. Local recolor checkpoint `c9ba419`
is above that observed remote head. The expanded CLI tests are a local-only
continuation; publishing them must preserve the separate recolor scope. Their
36 installed checks cover the unchanged earlier candidate DLL/CLI, not a fresh
build of the newer MCP guidance. Rebind and qualify any new packaged candidate.

Preview.3 is published at
https://github.com/GeorgeFejer91/wordplai/releases/tag/v0.1.0-preview.3.
Word was already closed when
the new open/save authorization was checked; the original documents were saved.
Installation and reopening then passed without changing their files. Browser
control recovered for the installed focus move, then native control was stopped
again with physical Escape during dialog testing. Await user-restored UI access
before claiming further browser scenarios. Do not infer shutdown from a
timeout or substitute public MCP checks for the named browser/installed surfaces.
