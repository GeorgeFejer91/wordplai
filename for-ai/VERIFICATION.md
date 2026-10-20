# Verification and readiness gates

Evidence must match the claim. Missing dependencies, credentials, hardware, or
runtime access produce `BLOCKED` or `NOT RUN`, never `VERIFIED`.

## Result vocabulary

- `VERIFIED`: the named check directly observed the claimed surface and passed.
- `PARTIAL`: some required evidence passed and the missing scope is named.
- `BLOCKED`: a concrete external or authority blocker prevented the check.
- `NOT RUN`: the check was intentionally not applicable or not attempted, with
  the reason stated.

## Gate 0: bootstrap readiness

From the repository root:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File for-ai/scripts/check-context.ps1 -ProjectRoot . -RequireRemote
git status --short
git rev-parse HEAD
git ls-remote origin refs/heads/main
```

Pass when the context checker succeeds, the intended tree is clean, and local
`HEAD` equals `origin/main`.

## Gate 1: task contract

Before implementation, name:

- the observable user outcome;
- affected product surface and owner;
- acceptance criteria;
- focused check that can fail for the requested behavior;
- broader checks required by affected boundaries;
- explicitly deferred work.

## Gate 2: focused change

Run the narrowest real check for the change. A syntax check proves syntax; a unit test proves
its tested logic; neither proves UI, deployment, hardware, performance, safety,
or scientific validity unless it directly observes that surface.

The standalone Windows component entry point is PowerShell 7:

```powershell
npm ci
npm run test:widget:state
./modules/word/test.ps1 -OutputDirectory artifacts/new-component-run
./modules/word/build.ps1 -Installer -NsisExecutable C:/tools/NSIS/makensis.exe
./modules/word/test-com-contract.ps1 -AssemblyPath artifacts/word/build/wrAIter.Word.dll -OutputDirectory artifacts/new-package-com
npm run test:site
```

Use fresh output directories. The component command starts no Word process;
`-NativeWord` and `-InstalledWord` are separate, opt-in acceptance surfaces.

## Gate 3: integrated readiness

Run proportionate build, test, lint, type, runtime, visual, device, security,
and compatibility checks for every affected boundary. Do not run an expensive
or irrelevant full matrix for a documentation-only edit.

## Cross-platform parity gate

For any change to a shared capability, update the contract in both repositories when needed and run the affected behavior checks independently on Word and Writer. A pass on one host does not qualify the other. A parity release is blocked while a shared capability is missing or unverified on either supported host; see [`CROSS-PLATFORM-COMPATIBILITY.md`](./CROSS-PLATFORM-COMPATIBILITY.md).

## Gate 4: publication

1. Review status and diff; preserve unrelated changes.
2. Confirm only intended paths are staged.
3. Confirm all required gates passed or are honestly reported.
4. Create one coherent commit under normal repository policy.
5. Push without force and without bypassing protection or secret scanning.
6. Verify the remote commit and required CI/deployment for that exact SHA.

## Handoff evidence

Report exact commands or observed surfaces, results, untested scope, commit SHA,
remote synchronization, CI/deployment state, and whether `for-ai/` changed.

## Standalone preview evidence — 2026-10-18

- `VERIFIED`: fresh `npm ci`, picker state checks, independent compilation,
  pinned official zrok acquisition, corrupt-archive refusal and four effective
  architecture negative controls. No predecessor checkout is a build input.
- `VERIFIED`: the new standalone component driver passed all 20 gates, including
  actual loopback transport (2,115 assertions plus copied mutation controls),
  real STA timeout/recovery, policy/recoloring, autosave doubles and 1,322 native
  WinForms checks with 20 renders. These do not establish an installed Word or
  live AI-client result. The driver's overall `PARTIAL` is intentional.
- After that run, packaging requirements and an early NSIS x64/Intel-AMD OS
  guard were added. The final installer was built from those post-change inputs;
  its exact DLL passed 536 COM ABI assertions and the effective dispatch-only
  negative control. The full component receipt remains bound to its earlier
  README/installer snapshot; no native C# or picker code changed between them.
- `VERIFIED`: all 40 native runtime files and 14 qualified picker/contract/asset
  files match their extraction SHA-256 in both working files and staged Git
  blobs. All 67 final build inputs and six packaged payload files matched the
  final build manifest after compilation; no source or payload drift was found.
- Final installer: `wordplai-setup.exe`, 23,874,671 bytes; SHA-256
  `923d6645b171e68d48930232b93eb738a4cb7530b16d198a9031b349e5aa6ccc`.
  DLL SHA-256: `0bd765fc5a23c02fe84815f2f347d39c9e2691d0eaf936ee4e1a60e7a562c19a`.
- `VERIFIED`: both unavailable and metadata-ready site branches separately
  passed 313 rendered checks in fresh headless Chrome with real bundled
  Pretext/fonts. Version, exact asset URL, checksum and size match the published
  preview. An anonymous HTTPS download returned the exact qualified installer.
- The first hosted Windows component run failed on a serialized-state string
  comparison: the expected and actual values matched recursively, but nested
  JSON property order differed. The test now canonicalizes object keys while
  preserving all values, fields and list order. Five effective comparison
  controls and all 355 policy assertions pass locally. No runtime code or
  installer input changed for this test correction; hosted requalification is
  tracked by the repository's required workflow result.
- `FAIL`: hosted Windows run `39126772674` passed policy, COM, dispatch,
  transport and autosave, then failed native UI containment: the Custom prompt
  GroupBox required 166px in a 115px table cell. The local native UI pass does
  not override this failure. The gate remains enabled; no runtime or test
  predicate was relaxed. Reproduction and acceptance are in `HANDOFF.md`.
- `VERIFIED`: Pages deployment `39126772649` and 86 independent live-browser
  checks at commit `073f0d5` observed both public routes, navigation, local
  fonts/Pretext, no overflow/errors and exact published release metadata.
- `NOT RUN`: installing this preview, clean-machine upgrade/uninstall, actual
  unsupported-OS refusal, Office bitness variants and physical native gestures.
  The installer is unsigned. Current foreground refusal, real WebGPT latency
  and host picker rendering remain open in [`HANDOFF.md`](HANDOFF.md).

Existing drivers' actual JSON receipts and logs are retained in ignored
`artifacts/`; they are the evidence source. The optional installed skill's
missing report helper was not run, and no helper-produced report is claimed.

## Full acceptance work — 2026-10-19

The primary goal now covers every applicable scenario in `HANDOFF.md`, including
installed and real-client surfaces. The earlier goal remains unfinished; these
results do not imply full acceptance.

- A fresh isolated Word workflow passed 209 assertions with zero drift across
  46 inputs at snapshot `2f7dfd0`. It covered independent native conversations,
  multiple/duplicate document names, closed/reopened choices, stale revisions,
  direct wording, comments, green protection, protected native units and native
  save/reopen. It used its own Word PID and disposable documents, not the
  installed add-in or a real WebGPT session.
- The UI candidate adds a readable no-fit state when the desktop or native
  maximum size cannot contain the current text metrics. Hidden editing controls
  cannot save through Enter; restoring space preserves the draft and returns
  the editor. The new bounded-window regression failed against the old code;
  1,368 native UI checks and 21 renders then passed locally. The final candidate,
  including context-marker recovery, passed all 20 component gates with 1,419
  native UI checks and 21 renders. Hosted run `39175387165` at `9ce27dc`
  subsequently passed all 20 gates, with 1,373 native UI checks and 21 renders
  on its smaller desktop. This qualifies the source fix, not preview.1's
  installed UI or physical gestures.
- `VERIFIED`: Chrome side-session control opened a fresh ChatGPT page, and its
  browser inventory exposed one profile. The API does not expose the Chrome
  toolbar avatar or profile selector. The ChatGPT page account icon showed
  orange initials, which does not confirm the yellow-figure Chrome profile the
  user identified. The requested profile and real WebGPT behavior/timing
  remain `BLOCKED`; the saved Wordplai connection previously returned
  `FORBIDDEN`. No retired connector or simulated client substitutes for that
  evidence. No account names, chat links or capability URLs are included here.

### Fresh isolated Word and browser evidence

All native drivers below used separately spawned, PID-bound Word instances and
synthetic documents. Each completed with unchanged captured inputs. The normal
Word process and its user documents stayed open. These checks do not qualify
the installed add-in or actual ChatGPT/Claude behavior. Counts overlap and must
not be added into a single total.

| Driver / focus | Observed result | Scope limit |
| --- | --- | --- |
| `test-native-suggestion-workflow.ps1` | PASS 209; independent conversations, duplicate document names, target retention, direct edits, native units, comments and reopen | Snapshot `2f7dfd0`; its original autosave path used manual coordinator pulses. |
| `test-native-suggestion-workflow.ps1 -DocumentChoiceOnly` | PASS 50; eight open documents, duplicate titles, conversation-local targets, inactive edits, copy detection, close/reopen rebinding, 64-conversation capacity and production inactive-document save/reopen; zero drift across 46 captured inputs | Fresh isolated Word process and disposable files; no installed add-in, normal Word session, physical interaction or real client. Receipt: `artifacts/word/qualification-20261019/native-choice-followup/manifest.json`. |
| `test-native-highlighting.ps1 -RecolorStress` | PASS 208; exact current paint, atomic refusals, stale identities, bounded inspection and recovery | Isolated source runtime. |
| `test-native-highlighting.ps1 -ColorStress` | PASS 46; neutral RGB, unchanged local-rule extension, ambiguity, batch limit and 82-marker reopen | No physical color picker interaction. |
| `test-autosave.ps1 -NativeWord` | PASS 19 native checks, component checks and four effective copied controls | Controlled clock/manual coordinator pulses. |
| `test-native-suggestion-workflow.ps1 -BackgroundSaveOnly -MutationCheck` | PASS 15; actual production timer saved an inactive selected document in 1,924.927ms, preserved its three green identities through reopen and left the active neighbor unchanged; disabled-timer copy failed the expected assertion | No manual Pulse or Save call after timer start; installed coordinator wiring still unrun. |
| `test-native-uncertain-outcomes.ps1 -MutationCheck` | PASS four native scenarios; three effective rollback/persistence controls | Copied fault seams, original runtime unchanged. |
| `test-native-precommit-conflict.ps1` | PASS; intervening native user edit retained, no staged paint/marker, stale witness revoked | Injected conflict in an owned fixture. |
| `test-native-accepted-changes.ps1` | PASS 63; complete acceptance, partial review, rejection, local rule changes and protected green wording | Isolated native review state. |
| `test-native-double-click.ps1 -NavigationOnly` | PASS 15; current range navigation and precise stale refusals | No physical double-click or native input event qualification. |
| `test-native-color-input.ps1 -RunNative` | PASS 11 cases, 180 outcome and 75 sampler-seam assertions; two unsafe-Undo copies failed the expected assertions | Injected input samples do not qualify the real foreground/Win32 input probe. |
| `test-native-highlighting.ps1 -DirectMode` | PASS 53; direct wording, rationale, green protection, exact fields/links and reopen | Isolated native source runtime. |
| `test-native-review.ps1` | PASS 337; marking, review, protected neighbors, stale refusals, saved-byte preservation, and stale cached prompt-marker recovery through a complete `context_changed` response and same-ID retry | Isolated synthetic Word process; installed package and public tunnel unrun. |
| `test-installed.ps1` | PASS 69; installed direct edits, rationale, autosave, save/reopen, protected text, revocation and refusal paths | Registered add-in `0.1.0.0`; its DLL hash matches neither packaged preview.1 nor preview.2. Public tunnel unrun. |
| Existing widget state and browser drivers | PASS; 32-document fixture, bounded pages, 12 cancellation/late-response cases, legacy and standard bridges, four widths and no browser errors | Synthetic host; no real ChatGPT iframe/bridge receipt. |
| `node site/test.mjs` | PASS 313; both routes, published preview.1 metadata, text geometry and readable fallbacks | Fresh local headless Chrome; live deployment is a separate observation. |

The focused `test-native-review.ps1 -StaleMarkerOnly` passed 19 assertions.
It observed the actual MCP refusal return `request_applied=false` and a complete
fresh context after repairing the still-visible prompt marker; the same request
identity then applied the exact wording, rationale comment and green protection.
The refusal preserved native state and saved DOCX bytes. A copied mutation with
the production `RefreshMarks` call removed failed at the intended
`marker_permission_unresolved` assertion and confirmed no native or saved-file
mutation; the original source hash stayed unchanged. Three earlier exploratory
receipts record harness assertions corrected against observed context shape and
marker multiplicity; they are retained as failed attempts, not passing evidence.
Final focused, mutation-control and full-suite receipts are in ignored
`artifacts/word/qualification-20261019/stale-marker-mcp-recovery-4/`,
`stale-marker-mcp-mutation-control-2/` and
`native-review-full-stale-marker/`, respectively.

The installed Word driver passed 69 checks against the currently registered
`wrAIter.Word` assembly (`0.1.0.0`, uninstall display version `0.1.0`). Its
registered DLL SHA-256 is
`1df260fbd4b4517b4aa7ebd422fedf42d952084e30bd325ddf8c7769cc5ed7ec`, which
matches neither packaged preview.1 (`0bd765fc...`) nor preview.2
(`0c691c88...`). This run is therefore an installed baseline, not evidence for
either release package. It used a synthetic saved/reopened document while six
existing Word documents remained open. The preservation audit retained all six
documents and their saved/revision/comment state, preserved local connection
state, closed the synthetic document, and restored the original active Word
window. The first cleanup wrapper attempt failed on COM by-reference argument
binding; the failed receipt is retained alongside the successful targeted
cleanup and final state receipt under ignored
`artifacts/word/qualification-20261019/installed-current-word-smoke/`.

The background timer's first new fixture failed because it prepared only one
of the three paragraphs required by the existing reopen assertions. The fixture
was corrected to prepare all three; the assertions and deadline were retained.
Its failing receipt and the old-code UI failures remain in ignored evidence.
Native workflow request-to-return timings ranged from about 0.82 to 5.12 seconds;
the focused timer fixture's three edits took about 0.93 to 1.38 seconds each.
These exclude upstream model delay and do not measure first-visible WebGPT edits.

### Next installer candidate

- `VERIFIED`: preview.2 builds with version metadata derived from `package.json`.
  Its 67 captured source inputs and six payload files match the build manifest.
  Its exact DLL passes 536 COM ABI assertions and the effective dispatch-only
  negative control. All 20 component gates pass with the candidate packaging
  inputs, including 1,419 native UI checks and 21 renders.
- Candidate `wordplai-setup.exe`: 23,874,145 bytes; SHA-256
  `8811ea09bd33f3add56549ef3bbcb6938f244d8b40121790dcb6f4d1515c43b7`.
  DLL SHA-256: `0c691c88a445d6d769654c23902267f64896f58c05ed83feae5c7a74008a3811`.
- `VERIFIED`: silent execution with Word open exits 2 and leaves all seven
  installed payload hashes, uninstall version and the original Word process
  unchanged. This observes the early refusal only, not installation.
- `VERIFIED`: the currently registered connected add-in passed 69 installed
  Word checks on a synthetic document. Its registered DLL hash
  (`1df260fbd4b4517b4aa7ebd422fedf42d952084e30bd325ddf8c7769cc5ed7ec`)
  matches neither packaged preview.1 nor preview.2, so this result qualifies
  that local installation only.
- `NOT RUN`: preview.2 installation/upgrade, installed runtime and UI,
  uninstall/clean-machine/Office-bitness/unsupported-OS variants, physical
  gestures and real WebGPT behavior/timing. Preview.2 remains a local candidate;
  the published release and stable download still identify preview.1.

Fresh private receipts are under ignored
`artifacts/word/qualification-20261019/`. Retain full source/harness manifests,
exact failures and native observations there; publish only this sanitized scope.

## Research handoff and current remote checks — 2026-10-19

- The source-linked test strategy and curated stress shortlist are in
  [`WORD-STRESS-TEST-RESEARCH.md`](WORD-STRESS-TEST-RESEARCH.md). It records
  primary platform/testing sources, three complementary qualification layers,
  high-risk identity and event-order scenarios, and proposed latency targets.
  The targets are experiment hypotheses only, not release criteria or
  measured results. The broader private source ledger and 42-row matrix remain
  ignored under `.for-ai-local/ui-stress-research/`; every row is
  `PROPOSED_NOT_RUN`.
- `VERIFIED`: hosted component workflow `39245471487` passed on
  `d2c95301134ee556d0ef126344a3dea41b8f5d0b`; workflow `39245802605` passed on
  `34c7f54f99fc98a9a528f14cd1f4b4e9f68b57c8`; workflow `39246178304` passed
  on `b74a1755a17e6f44b5df2276cfd03c260bc01329`; latest workflow `39246916923`
  passed on `2946052355b5b3e3b294e7d919218d8a0f37010d`, the then-current
  `main` tip. That run completed `npm ci`, widget state, Word component/transport
  checks without Office, retained native UI evidence and the For-AI structure
  gate. These component runs do not qualify installed Word, physical gestures,
  or real ChatGPT.
- `VERIFIED`: workflow `39291340698` passed on `a3161f2429b6c5c8e900eac89e48364c56efc1ba`.
  All steps passed, including 22-file native UI artifact upload via the pinned
  `actions/upload-artifact` v6.0.0 Node 24 action. The run log contains no
  deprecated Node 20 runtime warning.
- The required For-AI context gate initially reported a false remote failure
  under PowerShell 7.6.5 because `Select-Object` ran before `$LASTEXITCODE` was
  captured. The script now captures Git's output and exit code first, then
  filters the output. A rerun passed the required remote check; its expected
  working-tree warning remains until this change is committed.
- A fresh read-only Secret Tunnel availability call returned
  `UNAVAILABLE` / `invalid_mcp_response` after an HTTP 404 SSE response. This
  is distinct from the earlier saved Wordplai connector `FORBIDDEN`; cause is
  unconfirmed. The endpoint is deliberately omitted. Chrome side-session
  control can open a fresh ChatGPT page, but cannot inspect or select the
  browser profile; the requested yellow-figure profile remains unconfirmed,
  and no real-client check ran.
- At verification time, remote `main` resolved to the source SHA tested by
  workflow `39291340698`: `a3161f2429b6c5c8e900eac89e48364c56efc1ba`. The latest
  Pages workflow remains `39177712979` at
  `0b1cfc54fde09c4c852cbceb6656478a4db84657`; it still serves preview.1.
  Local documentation-only evidence updates do not alter product source,
  installer, release asset or site files.
- Follow-up commit `9da41535cf626e4860205f72eb9e7467b4963490` adds the focused
  stale-marker regression and fixes remote-status capture in the context gate.
  `VERIFIED`: hosted workflow `39294629741` passed all steps on that exact SHA.
  The local isolated Word review suite passed 337 assertions, its focused case
  passed 19, and the copied mutation control failed at the expected denial.
  The installed preview.2, physical gestures and real-client checks remain
  unrun; the open user Word document was left untouched.
- `VERIFIED`: hosted workflow `39297358857` passed all steps on documentation-
  only commit `e428820c9f29064c4b7f73fded401e2489e56da8`. It ran the widget,
  Word component/transport checks without Office, retained synthetic native UI
  evidence and passed the For-AI structure/safety gate. This does not qualify
  the installed preview.2, physical gestures or a real ChatGPT client.

## Artifact and public-state recheck — 2026-10-19

- `VERIFIED`: the local preview.2 installer and DLL still match their recorded
  SHA-256 values. `test-com-contract.ps1` against the exact preview.2 payload
  passed 536 ABI assertions; the dispatch-only mutant was refused before native
  invocation. Receipt: `artifacts/word/qualification-20261019/preview2-turn-recheck/com-contract/result.json`.
  This checks COM metadata and callbacks, not a running Word host.
- `VERIFIED`: the live `/wordplai/download/` route returned HTTP 200 and named
  preview.1 without mentioning preview.2. The public preview.1 installer
  returned HTTP 200; a fresh download measured 23,874,671 bytes and matched the
  published SHA-256 `923d6645b171e68d48930232b93eb738a4cb7530b16d198a9031b349e5aa6ccc`.
  The binary and sanitized response receipt are under ignored
  `artifacts/word/qualification-20261019/public-release-recheck-turn2/`.
- These package checks do not qualify preview.2 in installed Word. The open
  Word process, unconfirmed yellow-figure Chrome profile, real-client behavior
  and physical interaction remain outstanding.

## Active shipment qualification — 2026-10-20

All receipts in this section are under ignored
`artifacts/word/qualification-20261020/`. Counts overlap; do not sum them.

- `VERIFIED`: preview.2 installer SHA-256
  `8811ea09bd33f3add56549ef3bbcb6938f244d8b40121790dcb6f4d1515c43b7`;
  67 inputs and six payloads match. Silent upgrade with Word closed exited 0,
  six installed payloads matched and 99 existing profile files retained hashes.
  Installed DLL is `0c691c88a445d6d769654c23902267f64896f58c05ed83feae5c7a74008a3811`.
  See `installation/result.json`. Installed COM metadata passes 536 assertions
  with the effective dispatch-only negative control.
- `VERIFIED`: setup and uninstall refuse with Word open (exit 2), preserving eight
  installed files, display version and the original Word process. This qualifies
  refusal, not successful uninstall, clean-machine or other Office/OS variants.
- `VERIFIED`: widget state, synthetic browser picker checks across four widths
  and twelve late/cancel scenarios, and both site branches (313 checks) pass.
  These browser checks do not use ChatGPT's iframe. See `browser/results.json`.
- `FAIL`, retained: the first component/native run stopped at UI resources: GDI
  40→40, USER 38→42 after twenty cycles. Its native Word cases were not reached.
  The unchanged standalone UI recheck passed 1,419 checks (USER 43→43). Later
  passes do not establish the original failure's cause. See `components-native/`
  and `ui-resource-recheck/`.
- `VERIFIED`: all eleven isolated native suites exited 0: workflow (209), document
  choice (50), actual background timer save (15), review (337), recoloring (208),
  requested colors (46), direct wording (53), accepted legacy edits (63),
  precommit conflict, four uncertain-outcome scenarios and navigation (15).
  Effective copied fault controls remain retained. See `isolated/results.json`.
  These are PID-bound source fixtures, not installed/public/client evidence.
- `VERIFIED`: additional native autosave passes nineteen checks, component checks
  and effective controls. Color-input's good native variant passes and both
  unsafe-Undo mutants fail the exact preservation predicates, with unchanged
  captured inputs. Its original process nevertheless exited 1 because PowerShell
  inherited the last expected mutant exit. An explicit success exit now follows
  every existing assertion and integrity gate. The fresh native run exits 0,
  reports PASS with unchanged inputs, and retains both exact effective mutant
  failures. Reusing an existing evidence directory still exits 1. See
  `color-input-exit-fix/`, its process receipt and refusal receipt; retain the
  original `additional-native/` receipts.
- `VERIFIED`: installed preview.2 passes all 69 checks with Chrome foreground,
  retaining all six original documents' text, saved flag, comments, revisions and
  saved-file hashes. See `installed-candidate-background/result.json` and its
  preservation receipt. Earlier installed attempts remain failures: recovered
  cleanup barrier (normal Stop cleared it) and the intentional foreground Word
  guard. Do not call that guard resolved or transfer this pass to preview.3.
- `PARTIAL`: a physical double-click on the owned yellow passage opened the
  matching Make snappier dialog with direct guidance and no mode toggle; Cancel
  closed it. Physical draft modification/Save and complete native state assertions
  remain unrun. Raw screenshot stays private.
- `VERIFIED`: the revised overview assertion failed against the old production
  wording at the intended predicate (`connection-copy-regression-2/`). The single
  production text correction then passed all twenty component gates, including
  1,419 native UI checks/21 renders and GDI 40→40, USER 41→41. Receipt:
  `components-ui-copy-fix/result.json` (`components=PASS`; overall PARTIAL records
  the intentionally unrun installed/public surfaces).
- `VERIFIED`: preview.3 final build binds 67 unchanged inputs and six payloads.
  Installer: 23,868,622 bytes; SHA-256
  `a0f4965fa14e6980b64eeb79e6d60e497cb7d37b65b8103653ce05752e268605`.
  DLL: `5969541b4295760b2f78cab1e829f92b737a69a18aa958f90ea702d77e64b0f1`.
  Its exact DLL passes 536 ABI assertions and the effective dispatch-only control.
  See `preview3-final/verification.json` and `preview3-final/com/`. The first build
  attempt failed at an absent old NSIS path; final build used installed NSIS 3.11.
  Preview.3 is not installed or released. Its Word-open setup refusal passes
  (exit 2, all eight installed files/version/processes unchanged), recorded in
  `preview3-open-word/result.json`. User-controlled Word close is required.
- `PARTIAL`: a Chrome-extension WebGPT connection-only request attempted discovery
  and reported an internal tool failure. Reload temporarily showed an empty answer;
  a later native view showed the failed response. No picker cause is established.
  The saved endpoint differs from the candidate's new endpoint. The user approved
  an equivalent replacement connector for synthetic testing. It is now connected
  as `wordplai qualification`; its visible tool panel exposes exactly the same
  two read and two write tools. Old connectors were preserved.
- `FAIL`, retained: installed public transport briefly ran, then failed/revoked
  with no pending cleanup; HTTPS returned a tunnel-hosted 404. Separate exact
  installed TunnelProcess/pinned-zrok synthetic probes returned a browser warning
  interstitial, then API-client 502/timeouts. No warnings were bypassed, TLS checks
  disabled or permissions expanded. A later verbose probe passed: local exact
  response, external POST exact response at 13,011ms, child alive for another
  fifteen seconds, exact owned-share cleanup PASS. See
  `real-tunnel-fixture-api-client/` and `real-tunnel-fixture-verbose/`.
  One synthetic transport pass does not prove public Word/MCP reliability.
- `PARTIAL`: following a new installed connection start, a COM query found no
  add-in object. A later read-only probe finds nine open documents, four COM
  add-ins and Wordplai connected/exposed; status reports running/tunnelRunning.
  Cause of the earlier unavailable object remains unconfirmed. No document
  preservation after the earlier audit is inferred from this status probe.
- `VERIFIED`: actual public MCP initialization returned HTTP 200, application/json,
  a protocol version and session header (`public-mcp-recheck/result.json`). A fresh
  Chrome-extension ChatGPT Instant conversation, explicitly naming owned synthetic
  document A, executed its saved prompt. Native observation confirms the intended
  wording changed, one comment, zero revisions, exact protected emoji passage and
  Saved=true. The immutable saved DOCX snapshot then passes fourteen checks for
  every paragraph, native rationale, exact green paint and persisted protection;
  document B remains at its initial saved text/rules. See `webgpt-fixtures/saved-audit/`.
  The client displays 19 seconds for the whole task, not measured first-edit latency.
  DOM and screenshot confirm a visible answer; the AX tree omitted its text.
- `VERIFIED`: the post-WebGPT preservation audit, using the existing bounded Word
  message filter, confirms all six original documents unchanged in text, Saved,
  revisions, comments and saved-file hashes. The earlier unfiltered audit failed
  on an unavailable ActiveWindow and remains retained. See
  `after-webgpt-A-filtered-preservation.json`.
- `NOT RUN`: second independent live conversation, cold/reloaded host picker,
  live refusal/retry/revocation and first-visible-edit timings. B is prepared, but
  browser control detached after the second tab opened; its creation timeout did
  not mean it was absent. No duplicate writing request was submitted.
- `VERIFIED`, separate client scope: the reconnected app became callable from
  Codex. Through that actual public MCP connector, owned B was read completely;
  protected wording was refused with unchanged revision/text, authorized wording
  applied directly, exact replay added no duplicate rationale, and stale context
  returned `manuscript_changed` with `request_applied=false` and fresh context.
  The connector envelope labels this error `INVALID_ARGUMENT`; retain both layers.
  Saved B text, rationale, green protection and lack of revisions pass six checks;
  A's saved file remained byte-identical and all six originals passed preservation.
- `VERIFIED`, public MCP color scope: a fresh RGB created a neutral rule with empty
  prompt and two marks. A same-length wrong final quote refused the entire batch,
  preserving text, revision, rules and markers. Neutral color did not authorize
  wording. Reusing the existing neutral identity preserved its rule; immediate
  exact replay returned the same marker without duplication. A prior replay test
  incorrectly expected success after an intervening native refusal had revoked
  authority; the actual safe stale/revoked refusal is retained, not relabeled as
  a successful replay. The source explicitly binds receipts to their authority
  epoch/session. Seven saved-DOCX checks pass for exact text/paint, six total marks,
  three neutral marks, one empty-prompt neutral rule, one rationale and no revisions.
  All six originals again pass preservation. Receipts: `public-mcp-complete-part-*-private.json`,
  `webgpt-fixtures/public-mcp-B-saved-result.json`,
  `webgpt-fixtures/public-neutral-saved-result.json` and
  `after-public-neutral-preservation.json`. This supplements the Chrome result;
  it does not qualify a second WebGPT conversation or host picker rendering.
- `VERIFIED`: hosted component workflow `39425758523` and Pages `39425758498`
  both passed on `40ba9568507274815b5cc5ae49c6003f4313223e`. A fresh public Download
  HTTP 200 still matches preview.1's version, asset URL and checksum and does not
  advertise preview.3 (`pages-after-source-push.json`). No new installer release
  was created. The context gate with `-RequireRemote` passes on the pushed clean
  checkpoint; later documentation checkpoints require their own remote recheck.
- `FAIL`, retained: hosted `39427161161` on `b79688d` stopped at recolor state
  preservation. Parsing its actual expected/actual JSON with the existing policy
  comparer confirms identical values/fields/list order and different object-key
  ordering (`hosted-recolor-comparison-diagnosis.json`). The existing canonical
  comparer is now shared by the two policy suites as `PolicyStateJson.cs`; the
  driver compiles that test helper into both. Production serialization and the
  installer remain unchanged. The five strict comparison controls still pass;
  a copied scalar-losing mutant fails the exact changed-identity predicate.
  The first partial repair missed one reopened-state comparison and failed;
  that receipt remains in `components-state-comparison/`. After correcting both
  sides, all twenty component gates pass, including 168 recolor assertions,
  1,419 native UI checks and 21 renders. See `components-state-comparison-final/`,
  `hosted-recolor-comparison-fixed.json` and `state-json-negative-control/`.
  The final candidate's 67 inputs and six payloads still match, with unchanged
  installer SHA-256 (`preview3-final-current-binding.json`). Hosted verification
  `39428499738` passed on the exact published correction `7034ff0`.
- `VERIFIED`, preview.3 installed lifecycle: upgrade from preview.2 and reinstall
  exit 0, with all six payloads, manifest, version and both registration views
  correct. Uninstall removes the payload/uninstaller and all eight checked COM,
  Word add-in and product keys across both registry views. All 107 existing
  profile files remain unchanged, with none added during this lifecycle. See
  `preview3-lifecycle/result.json` and `preview3-unregister/result.json`. This
  covers this existing-profile host, not a clean machine or other Office/OS hosts.
- `VERIFIED`, original preservation: Word was already closed and all six originals
  saved when the user's new open/save authorization was checked. Fresh installed
  Word reopened all six with unchanged native text, comments, revisions, Saved
  state and file hashes; the preview.3 add-in is connected/exposed with its exact
  DLL hash. See `preview3-reopen/result.json`. The first helper had a parser error
  before execution and remains retained.
- `FAIL`, retained: preview.3's first installed runtime attempt stopped after
  three checks at recovered connection cleanup, with all six originals unchanged
  (`preview3-installed-preservation.json`). Normal Stop subsequently cleared
  cleanup and busy state. The new Word window held foreground focus; normal
  Chrome activation subsequently moved focus without removing the guard.
- `VERIFIED`, preview.3 installed runtime: the fresh
  `preview3-installed-background/result.json` passes all 69 checks, including
  direct editing, native rationale/green protection, exact replay, stale feedback
  refusal, Undo/Redo, Discard/Reject, timer autosave, revocation and saved/reopened
  markers. The matching preservation receipt passes for all six originals with
  exact DLL SHA-256 `5969541b4295760b2f78cab1e829f92b737a69a18aa958f90ea702d77e64b0f1`.
  This driver uses actual installed Word/MCP loopback, not the public tunnel.
- `PARTIAL`, physical preview.3 UI: a real double-click on owned synthetic C's
  yellow prompt passage opens the correct color-rule dialog with direct-edit
  guidance and no mode toggle. Its accessibility tree is empty. The dialog image
  did not change after attempted field input, but the later native audit finds
  C's body replaced with the disposable draft string and automatically saved.
  This is a failed automation interaction, not a verified dialog draft edit or
  an AI write. No dialog Save occurred; C is invalid for further writing tests.
  The attempted Cancel was stopped by the user's physical Escape, so
  draft/Cancel/Save remains unqualified. Raw diagnosis and the initial preparation
  precondition failure are retained under `preview3-webgpt-fixtures/`. The first
  audit also found that legacy PowerShell's default JSON file encoding lost the
  baseline emoji; use explicit UTF-8 or the source-defined Unicode expectation
  for future fixture audits. Chrome control recovered to the original A chat
  before that stop; no new WebGPT writing request was submitted.
- `VERIFIED`, post-UI original preservation: the bounded native audit again
  passes for all six originals, including exact text, comments, revisions, Saved
  state and file hashes (`preview3-after-native-ui-preservation.json`). The
  disposable C incident does not authorize claims about dialog input routing.
- `VERIFIED`, preview.3 public restart/revocation: the installed post-restart
  HTTPS endpoint initializes MCP with protocol `2025-03-26`, HTTP 200 in 629ms.
  Normal Stop reaches stopped with no cleanup pending; initialization at the old
  capability URL then returns HTTP 404 in 126ms. Sharing is stopped. This is one
  transport/revocation case, not sustained reliability or live WebGPT editing.
  See `preview3-public-revocation/result.json`.
- `VERIFIED`, Windows preview shipment: prerelease `v0.1.0-preview.3` targets
  `6c455dbc1995e8f722b8d3b6881c7d9a135125f1`, whose hosted verification
  `39431955158` passed. Its three uploaded assets match their local SHA-256 and
  size. Anonymous HTTPS installer download returns HTTP 200, 23,868,622 bytes,
  SHA-256 `a0f4965fa14e6980b64eeb79e6d60e497cb7d37b65b8103653ce05752e268605`.
  The source/payload manifest SHA-256 is
  `611af95264d1c56221675d32baa9275d92ee6dbbd90031c86b2800e680de62c9`.
  All 67 inputs and six payloads remain exact after installed qualification.
  The release is unsigned and explicitly lists the incomplete physical,
  preview.3 WebGPT/picker/timing, clean-machine and Mac scopes. Publication does
  not complete overall acceptance. See `preview3-release/` and
  `preview3-after-installed-binding.json`; verify subsequent site deployment
  separately.
- `VERIFIED`, preview.3 site source: the existing rendered matrix passes 313
  checks after updating release copy and the published asset's exact metadata.
  Minimum-width Description and enlarged-text Download captures were inspected,
  followed by the final desktop Download with preview.3's checksum and limits.
  Locked Pretext/fonts, keyboard navigation and readable fallbacks remain covered.
  See `preview3-release/site-published-metadata-results.json`; this local render
  does not establish Pages deployment or the ChatGPT picker.
- `VERIFIED`, preview.3 deployment: hosted verification `39433032039` and Pages
  `39433031883` both pass on `25a992f0a1a838c7b262d25a75e4f933f288b76c`.
  Anonymous Description, Download and release JSON return HTTP 200 and match the
  locally tested output exactly. All twelve deployed build-manifest inputs match
  that published Git commit. The first raw local-manifest equality assertion
  failed: five text inputs differ only by Windows CRLF versus published LF.
  Both the failure and explicit normalized-byte/Git-blob diagnosis are retained;
  no file content or check was silently discarded. Anonymous manifest/checksum
  assets also match the release staging files. See
  `preview3-release/deployed-pages-result.json`, `deployed-manifest-diagnosis.json`,
  `deployed-line-endings-diagnosis.json` and `anonymous-metadata-result.json`.
- `DEFERRED BY USER`: Mac implementation, installer and installed Word/client
  qualification. The user has no Mac and requested a future-agent handoff;
  `MACOS-COMPATIBILITY.md` now specifies starting context, capability feasibility,
  artifact/host acceptance and handback. No Mac support claim is made.

## Windows CLI qualification — 2026-10-20

Private evidence root: `artifacts/word/cli-qualification/`. CLI commands use local
COM authority for rules/markers and the existing four MCP tools for AI edits.
No desktop input or new remote authority is introduced.

- `VERIFIED`, source-native: `test-native-review.ps1 -CliOnly` passes 33
  assertions (`native-final/`): exact/closed/duplicate/inactive targets, stale
  snapshots, Unicode boundaries, wrong quotes, protection, marker removal,
  unchanged neighboring saved bytes and selection/foreground preservation.
  The original red run fails at the missing automation operation as intended.
- `VERIFIED`, CLI client: 14 checks including the effective input-limit copied
  control pass against the packaged executable (`preview4/cli-exit-fixed/`).
  Its driver explicitly exits zero after expected negative controls; the earlier
  inherited nonzero wrapper exit remains retained. Picker state doubles pass.
- `VERIFIED`, components: the first complete CLI component run passes all 21
  gates (`components/`), including 1,419 native UI checks and 21 renders. A later
  full run fails at the unchanged foreign-share cleanup timeout; its cause is
  unconfirmed. A scoped transport recheck passes 2,115 assertions and effective
  controls with all fifteen transport inputs unchanged. That scratch wrapper
  inherited an expected negative-control exit; do not call the later full run a
  pass. Final hosted verification is recorded separately after publication.
- `VERIFIED`, preview.4 candidate: 71 inputs/eight payloads, exact DLL COM ABI
  536 assertions and dispatch-only control, upgrade from preview.3, complete
  uninstall/registration removal and reinstall. All 112 profile files remain
  unchanged, none added. Installer SHA-256:
  `ed4025a3b161c995cb843aca4c5f5900cd2cbebbed329b3c9ac355e7dfc9b130`
  (23,888,654 bytes); DLL:
  `5c10abcb79071d97d166dcb64132eb4f40da44911511fcffbec787f739bff5f0`;
  CLI: `381df69f63ff2e6e963a73c89ef09b943423bf075a5627b3667e38292db3ff0b`.
- `VERIFIED`, actual installed CLI: `test-installed-cli.ps1 -WordProcessId PID`
  passes all 16 checks (`installed-existing-final/`), including local rule/range
  marking, stale refusal, native active-document/selection/foreground equality,
  real loopback sessions and complete-read witness, direct wording, exact replay
  with one rationale and no tracked revisions, unchanged neighbor and all six
  pre-existing documents. No desktop input; public tunnel and WebGPT not tested.
- `VERIFIED`, preservation: save/close and reopen audits preserve all six
  originals, their exact text/comments/revisions/Saved state and file hashes
  (`headless-test-close/`, `reopen-after-cli/`); the post-test
  `installed-final-originals-result.json` again verifies all six native states
  and file hashes. The installed add-in is connected
  with the candidate DLL. Local sharing started by the driver is stopped normally.
- `FAIL`/`PARTIAL`, retained installed attempts: a foreground-Word run skips the
  positive edit; second-instance attempts refuse the single-owner profile;
  a mixed-format fixture safely refuses without wording/revisions/comments.
  The corrected fixture uses the existing installed suite's coherent font setup.
  A later 15-check successful prefix fails preservation of a transient EndNote
  startup template; a diagnostic captures that template as the original. Another
  diagnostic fails foreground equality, without proven cause. Startup-readiness
  and `/a` attempts fail before creating fixtures; a driver-edit error is repaired
  before the final passing run. The optional installed-process launcher was
  removed, not qualified. Normal installer registration restoration was used;
  no Office security, profile ownership or production guard was weakened.

Preview.4 is a local installed candidate; the public release/site remain preview.3.
Concurrent recoloring source/test/document edits appeared during final review;
preserve them outside the CLI checkpoint and compare the candidate to the staged
Git inputs. Do not claim the whole working tree matches that installer. Physical
UI, real WebGPT, clean-machine and Mac acceptance remain open.

Raw browser captures, endpoint capabilities and document-preservation snapshots
remain private. Retain exploratory harness failures. Native source checks run
serially; do not change captured inputs while running. Native files intentionally
preserve CRLF under `.gitattributes`; review with Git's `cr-at-eol` interpretation.
