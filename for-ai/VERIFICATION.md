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
