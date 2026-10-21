# Verification and readiness gates

Evidence must match the claim. Missing dependencies, credentials, hardware, or
runtime access produce `BLOCKED` or `NOT RUN`, never `VERIFIED`.

## Result vocabulary

- `VERIFIED`: the named check directly observed the claimed surface and passed.
- `PARTIAL`: some required evidence passed and the missing scope is named.
- `BLOCKED`: a concrete external or authority blocker prevented the check.
- `NOT RUN`: the check was intentionally not applicable or not attempted, with
  the reason stated.

## Comment author and installed preview.9 — 2026-10-21

- `VERIFIED`: Setup explicitly saves a local comment author for new rationale
  comments, defaults to wordplai, retains Unicode and leaves existing comments
  and the global Word user name unchanged. Five-field profiles load their old
  default. Empty/control/invalid-Unicode/overlength names refuse; an external-file
  conflict preserves both the effective name and the external saved state.
  Restarted profile objects retain the exact saved Unicode name. The transport
  suite passes 2,146 assertions and fifteen schema cases, with its existing
  copied fault controls. The initial regression cannot compile against the old
  missing preference API; the native behavioral fault control below is stronger.
- `VERIFIED`: `modules/word/test.ps1 -OutputDirectory
  artifacts/word/comment-author/components-final` passes all 21 component gates.
  Its UI gate passes 1,452 checks/21 renders. Subsequent focused UI checks add
  native traversal from the name field to Save, reopen/discard of the draft and
  a width/height matrix, passing 1,738 checks/21 renders in `ui-matrix/renders/ui-report.json`.
  The first enlarged-text run fails on the vertically clipped Save row. The
  corrected native row pairs input and Save; minimum 640×700, tall/narrow,
  wide/short, expanded and enlarged-text layouts retain visible controls without
  panel scrollbars. This is native WinForms measurement, not HTML/Pretext or
  installed physical-keyboard evidence.
- `VERIFIED`: `test-native-highlighting.ps1 -DirectMode -MutationCheck` passes
  56 isolated Word assertions. A chosen Unicode author appears on the actual
  native comment and in saved/reopened DOCX; wording, native units, protection,
  Undo/Redo and replay checks pass. A copied hardcoded-wordplai implementation
  fails the exact saved-author predicate. Initial harness compile/command errors
  remain retained; they do not replace these actual Word results.
- `VERIFIED`: exact candidate COM 536/CLI fifteen and their copied controls pass.
  NSIS 3.13 portable is retrieved from the official project and checked against
  its HTTPS release RSS size/MD5; archive SHA-256 and tool bytes remain private.
  The documented C:/tools compiler path is absent on this PC. Early download
  attempts return HTML and fail the checksum gate before execution; the actual
  release archive passes. The 23,886,825-byte unsigned preview.9 installer SHA-256
  is `fe4a106e1ec5d66980ccde64cd6fa51bc474686f95743b18a13853315d57863d`.
  All 72 source inputs/eight payloads match. Word-open setup refuses with exit 2
  and preserves installed bytes. Normal stopped Word exits before upgrade;
  both registration views, manifest, version and 150 private profile files pass.
- `VERIFIED`: the exact seven saved documents reopen unchanged with their active
  document and selection restored. Hidden startup first exposes only an empty,
  saved EndNote startup template, so the initial ROT observer expires and the
  initial expected-document observer refuses. PID-bound recovery reuses the
  same process, opens the exact saved originals through COM, and normally closes
  only that verified empty startup template. No original is discarded or edited.
- `PARTIAL`, first installed CLI run: eighteen checks pass but Word foreground
  prevents positive writing. `VERIFIED`, separate minimized/background run:
  `test-installed-cli.ps1 -IsolationCases -CommentAuthorName 'Review Editor'`
  passes 45 checks, positive MCP writes included, with zero unqualified foreground
  checks. Both documents' native and timer-saved rationale comments carry the
  exact chosen author. Replay, cross-client isolation, revocation/recovery,
  closed-target and original-preservation checks pass. The original author
  preference is restored; no mouse/keyboard input is used by this driver.
- `VERIFIED`: public access is running in installed Word PID 39856. A plain
  non-browser HTTPS probe initializes with HTTP 200, discovers exactly four
  existing tools and reads the exact packaged 58,378-byte picker, in 1,659ms
  combined. This is transport/resource evidence, not renderer/edit latency.
  Initial probes receive 400 with an unsupported underscore header and warning
  HTML with the default browser-like user agent; retain those failed probes.
  Capability/session addresses are private. Final seven-file native/disk
  preservation passes, with all originals saved and unchanged.
- `VERIFIED`: picker state checks and 313 site checks pass via their direct Node
  entry points (npm is absent from this shell). No site source/asset is changed.
  The pinned picker HTML is unchanged; its build manifest reflects the versioned
  lockfile. `git -c core.whitespace=cr-at-eol diff --check` honors the repository's
  byte-preserved native files; ordinary diff check treats their CRLF as trailing
  whitespace. The context checker passes with existing ledger-size warnings.
- `NOT RUN`: preview.9 real Chrome host rendering/wire, installed physical
  keyboard/custom typing, normal first-visible edit timing, clean-machine and
  future Mac qualification. No enabled Chrome/native control surface is present
  in this session; Windows Sandbox is absent and a separate clean host is not
  identified. Writer author-preference parity is unimplemented/unverified.
  The full release goal stays open. Raw receipts are ignored under
  `artifacts/word/comment-author/`; no preview.9 publication or hosted CI is claimed.

## Real client/native continuation and overlap source correction — 2026-10-21

- `FAIL`, fresh warm-cache Chrome conversation: first picker rendering and the
  ordinary assistant response report load failure. `VERIFIED`, reload recovery:
  the picker renders seven choices, selecting owned G succeeds, and the automatic
  continuation reports a complete read with protection taking precedence. The
  previous H conversation then reports its retained H target on a new unnamed
  complete-read request; a reciprocal G request retains G. No wording/color/rule
  change is requested. The final seven-open/six-saved-original audit passes.
  Resource reads/results/completed writes remain two throughout, with zero
  RPC/resource/Origin errors; HTTP requests/refusals increase separately. This
  establishes a cached first-render failure, not its cause or captured bridge wire.
- `PARTIAL`, browser keyboard: Tab through the nested picker times out after
  focusing H; the focus observation remains on H. No successful Tab/Enter selection
  is claimed. The subsequent ordinary visible G click succeeds. Private browser
  traces and images are in `artifacts/word/resource-diagnostics/preview7-warm/`.
- `VERIFIED`, installed preview.7 physical mouse path: ribbon double-click and
  marked-text double-click open the exact owned rule. A preset checkbox changes
  the draft; 21 checks confirm no commit or saved-byte change. Cancel closes it
  and independently passes the same 21 checks. Reopening shows the preset absent.
  Selecting it again and clicking Save closes the dialog, commits the exact preset
  plus original custom prompt, and reaches ordinary timer save. A further 21 checks
  preserve text, mark identities/ranges/quotes, native RGB, annotation counts,
  unrelated rules and all seven pre-existing documents/six saved originals.
- `VERIFIED`: ten independent actual-DOCX checks correlate the physically saved
  rule with CLI state, exact text/paragraphs, all saved mark fields, RGB, no tracked
  revisions/comments and no review proposals. A copied persisted policy omitting
  the selected preset fails the exact saved-prompt assertion. The initial fixture
  constructor expectation omitted Word's extra retained terminal paragraph; its
  refusal is kept. Exact observed paragraph verification succeeds without changing
  wording. Retain the initial shared-file hash and by-reference Range audit errors;
  corrected read-only observers use shared read access and explicit COM references.
- `NOT RUN`, custom dialog typing: the supported helper returns only the Word
  owner window. Mouse input on its related dialog screenshot works, but Ctrl+A
  after prompt focus selects the document body. Testing stops before text entry.
  Alt/F10 attempts do not establish ribbon keyboard navigation. These checks do
  not certify custom typing, keyboard access, different desktop/DPI configurations
  or the earlier invalid C fixture. Private physical evidence is under
  `artifacts/word/native-ui-preview7/`.
- `FAIL`, installed overlap chooser: two exact marker identities render four
  selectable rows while the pager reports two. Selecting and opening the repeated
  third row does nothing. Closing the chooser preserves the owned policy exactly.
  Only the new UI fixture is then normally closed, with saved bytes unchanged;
  the final original seven-document set and six known saved files remain exact.
- `VERIFIED`, source regression/correction: the native ListBox row count is now
  checked against managed Items before palette refresh can mask the defect. Prior
  source reproduces native=4/managed=2. A private trace observes synchronous renders
  during WinForms creation/font/resize processing. The shared list coalesces these
  layout callbacks on its UI control after those operations; paging and selection
  remain unchanged. Final focused UI passes 1,433 checks and 21 renders, including
  exact selection after resize/font changes. The final test with the original list
  fails the intended native-count predicate; both runs have zero input drift.
  Keep the earlier trace-preparation line-ending refusal and first driver's stderr
  capture failure separate from the executed regression. Evidence is in
  `artifacts/word/overlap-list-fix/`. The full source component driver passes all
  twenty-one gates, including the 1,433 native UI checks/21 renders, existing
  transport and safety fault controls, COM 536 and CLI fifteen. Its overall
  `PARTIAL` retains installed/public-client scope. Source checkpoint `8044b36`
  passes hosted verification `39515532203`; the following distinct preview.8
  qualifies the installed correction. Preview.7's original failure is retained.

## Installed preview.8 CLI and native correction — 2026-10-21

- `VERIFIED`: 72 source inputs/eight payloads match the distinct package;
  exact-package COM 536 and CLI fifteen checks and effective fault controls pass.
  Installer size is 23,885,366 bytes, SHA-256
  `E0F8D9456798D8507021A893E48421ACE9C8AE7EA67CC3104B152AB75A8DE423`.
  Installed DLL SHA-256 is
  `3E8A349F70CD03ED984E29C4B6B6FAE65BB0B2A722B3F56FF51D2161FECD3ACC`;
  installed CLI SHA-256 is
  `91D185C7476BB87A29D0FA6EE9C79A673A1EC6C28C4A43836CBC28F77EE458C6`.
- `VERIFIED`: Word-open installer refusal preserves eight installed files.
  Settled normal Stop precedes Word exit and upgrade. Payloads, version/manifest
  and both registration views match; 143 profile files remain unchanged, none
  added. Seven documents reopen with exact saved bytes, native text, annotations
  and Saved state. Public preview.5 assets remain unchanged.
- `VERIFIED`: installed CLI passes 42 checks, positive MCP writing included,
  with zero unqualified foreground comparisons and no mouse/keyboard input.
  Coverage includes separate clients, complete-read witnesses, cross-client
  refusal, replay, revocation/recovery, local-only background timer autosave and
  closed targets. Nine installed numeric-diagnostics checks pass independently.
- `VERIFIED`, actual native mouse interaction: corrected overlap chooser shows
  exactly two rows for two records. Selecting the second opens its exact local
  protective rule with Preserve checked and wording controls disabled. Cancel
  and Close leave policy and wording exact. Only the owned UI fixture closes;
  all seven prior open documents/six known saved originals remain unchanged.
- `FAIL`, equivalent four-tool Chrome app: cold and reloaded picker both report
  display/response-load failure. Naming fresh owned I once completes its real
  WebGPT edit. Naming J once produces a response-load failure, including after
  reload, while Word applies and saves the edit once. A read-only follow-up
  displays applied=true/replayed=false and the existing result; no writing request
  is resent. Both actual saved DOCX files pass fifteen checks each: exact displayed
  Unicode wording, protection/ranges/RGB, unchanged rules, one applied receipt,
  one exact anchored rationale, no tracked revisions and other-document isolation.
  The copied missing-protection control fails as intended. These are displayed
  client/native correlations, not captured raw client wire.
- `VERIFIED`: the read-only report changes neither saved bytes, native text and
  annotation counts nor complete CLI state. Only owned I/J close normally with
  byte-exact saved files. Final seven-open/six-saved-original audit passes.
  Retain the initial audit's PowerShell 5 UTF-8 decoding mismatch (independent
  JSON parsing confirms identical values), and the earlier UI-fixture cleanup
  path-substitution failure; both stop before closing their fixtures. Corrected
  explicit UTF-8/path observers pass. Raw receipts remain private under
  `artifacts/word/overlap-list-fix/preview8/`. At that runtime checkpoint the
  candidate was local and unpublished; its later release is recorded below.
  Reliable host rendering, keyboard/custom typing, raw wire,
  first-visible latency, clean-machine and Mac scopes remain open.
- `VERIFIED`: version/evidence checkpoint `ae583e7` is pushed and passes hosted
  verification `39521101640` and Pages `39521101662`. Local rendering passes 313
  checks. Three anonymous page/release files match tested bytes; all twelve
  deployed source hashes match exact Git blobs. The copied changed-lock-hash
  control fails its intended predicate. The public release remains preview.5.

## Preview.8 publication — 2026-10-21

- `VERIFIED`: the candidate's 72 source inputs, eight packaged/installed payloads
  and exact installer are rechecked before publication. Version/evidence target
  `994f03473bade689e6ad5f58482ecb88e1e19faf` passes hosted verification
  `39521437016`. Public prerelease `v0.1.0-preview.8` targets that exact commit;
  the remote tag resolves correctly. All three assets return anonymous HTTP 200
  and match local tested bytes, sizes and GitHub SHA-256 digests. Attached sums
  independently match both installer and manifest. Preview.5's assets are unchanged.
- Public installer: 23,885,366 bytes, SHA-256
  `e0f8d9456798d8507021a893e48421ace9c8ae7ea67cc3104b152ab75a8de423`.
  Public manifest SHA-256:
  `9b0e7a6966dc096e02fb06e79a2e3358bd09fc9f65d13f03ddb2eb47ec52114b`;
  checksum file SHA-256:
  `cbc41fd43d58e647b6e260bcfe36725aefcd488a23774ea608b08cae3dc4b5fa`.
  Raw public API/download receipts remain private under
  `artifacts/word/preview8-release/`. No clean-machine or Mac support is inferred.
- The release notes and updated site/README explicitly retain browser display/
  response failures (including a committed edit behind a failed response),
  foreground refusal, native keyboard/custom typing, latency, additional host
  configurations, clean-machine and future Mac gaps. Physical preview.7 preset
  results retain their version; preview.8 separately qualifies corrected overlap.
  The actual-metadata site's verification is recorded below.
- `VERIFIED`: actual-release site checkpoint `e23d1e8` passes hosted verification
  `39522289127` and Pages `39522289143`. Local rendered verification passes 313
  checks across both routes, sizes, enlarged/localized/spacing text and readable
  fallbacks. Three anonymous page/release files match tested bytes, and all twelve
  deployed source hashes match exact Git blobs; the changed-lock-hash control is
  effective. Site version/time/URL/hash/size independently match published GitHub
  asset metadata. No host widget success is inferred from the site result.
- `VERIFIED`: the post-publication audit preserves all seven open documents and
  six known saved originals with zero native/file changes. Its private receipt is
  `overlap-list-fix/preview8/post-release-preservation-result.json`. Full release/
  deployment receipts and rendered output remain under `preview8-release/`.

## Preview.8 real-client closed target and reopen recovery — 2026-10-21

- `VERIFIED`, scoped client/native correlation: one explicitly requested real
  Chrome read of closed owned J reports `document_closed`, no fallback/retry and
  `request_applied=false`. All displayed code/message/details exactly match a
  separate actual CLI MCP refusal, whose text/structuredContent agree and whose
  outer tool result has `isError=true`. The chat reports `isError` absent and an
  outer `INVALID_ARGUMENT`; this is not captured raw host wire. Fifteen checks
  pass, including exact inventory/active target, seven-open/six-saved-original
  preservation and unchanged saved I/J files. A copied client-code mismatch fails.
- `VERIFIED`: only owned J reopens hidden, retaining identity, rules/marks,
  pending review state, wording, saved bytes and one exact rationale. Explicit
  real-chat complete reading recovers that target; a follow-up client-reported
  offset-only call returns the same J while another Word document remains active.
  Thirteen checks correlate both actual rendered tables with native/CLI state,
  complete flags, new sharing authority and exact rules/mark fields. Wording,
  rationale content/anchor and saved bytes remain exact with no new revisions.
  Normal close restores seven documents; all originals remain unchanged. A copied
  displayed-identity mismatch fails its intended predicate. Word lifecycle uses
  no mouse/keyboard input. Real request arguments/call counts remain client reports.
- The original J edit response now renders after reopening its chat; its earlier
  first/reload failures remain retained. Inspected supported Chrome capabilities
  expose console/page assets, not network tracing; response actions expose no raw
  result view. The only two captured warnings concern Statsig cache writes and do
  not identify a cause. During the closed-target chat, HTTP requests/refusals rise
  by four/one; resource/RPC/Origin errors remain zero, with no resource read/write.
  Aggregate counters do not prove the exact real-client call sequence.
- Evidence is private under `artifacts/word/webgpt-preview8-closed-target/`.
  Runtime, package, endpoint and published preview.8 assets are unchanged.
  Raw live client wire/bridge, reliable rendering, keyboard, visible timing,
  clean-machine and future Mac scopes remain open.

## Preview.8 visible edit with host approval delay — 2026-10-21

- `VERIFIED`, scoped native outcome: one real Chrome/Instant named-document
  request edits fresh owned K to its exact saved local prompt, retaining the
  emoji and all unmarked/protected wording. Fourteen independent timer-saved
  DOCX checks verify rule/marker identities, UTF-16 rebasing, exact protective
  RGB, one applied receipt and one anchored rationale, with zero tracked
  revisions. Native/CLI correlation passes; removing protection from the audit
  input fails the intended predicate. No writing request is resubmitted.
- `PARTIAL`, visible timing: Word stays visible in the background with an owned
  disposable focus window foreground. The first host approval click times out;
  native wording and the same pending approval remain unchanged. A later semantic
  Allow once succeeds without persistent permission changes. Retained JPEG frames
  bracket changed pixels at 163.083–500.306 seconds after Send, including approval
  and operator delay; changed wording/comment are observed by 5.249 seconds after
  the acknowledged approval click. These are coarse workflow bounds, not normal
  latency or an exact approval-to-edit measurement. The sentence is partly clipped
  in those frames; exact complete wording is verified independently above.
- `FAIL`, host answer: the screenshot reports "This response couldn't load"
  although accessibility says Response complete. Frames showing the native edit
  retain old UIA document text; cause/duration are unqualified. One applied native
  edit persists despite the failed host answer. Raw host wire is not captured.
- `VERIFIED`: normal close of owned K preserves its saved bytes; the owned focus
  window closes and Chrome returns to its original minimized state. All seven
  prior open documents and six known saved originals remain unchanged. Retain
  the initial UTF-8-without-BOM observer failure and corrected pass separately.
  An initial capture with mismatched pixels is excluded from timing evidence.
  Private packet: `artifacts/word/preview8-visible-timing/qualification-result.json`.
  Runtime, package and public assets are unchanged. Reliable rendering, normal
  timing, keyboard/custom typing, clean-machine and future Mac scopes remain open.

## Preview.8 installed CLI maximum batches — 2026-10-21

- `VERIFIED`: the exact published DLL/CLI pass eighteen installed checks through
  actual local MCP, without mouse/keyboard input. Batches of 1, 9 and 64 ranges
  return exact marker identities/ranges and native RGB. A 65-range batch and a
  duplicate-ID recolor refuse without native, policy or saved-byte changes. A
  wrong final quote also preserves every native/saved field and rule/marker,
  while explicitly revoking sharing. A fresh complete read restores authority.
  The original assertion requiring sharing to remain true fails and is retained;
  the corrected assertion requires that revocation and forbids every other change.
- `VERIFIED`: recoloring 64 selected markers retires exactly those identities,
  preserving ten earlier marks and the protected ending. Exact replay returns the
  same current/retired identities and preserves full CLI, native XML and saved
  bytes. Eighteen separate saved-DOCX/receipt checks verify all 75 native bookmark
  ranges, exact Unicode/paragraphs/RGB, unchanged old rules, two neutral new colors
  and no tracked wording revisions/comments. A missing-marker control fails the
  intended persisted-count predicate. The 64-range mark/recolor CLI round trips
  take 2,235/3,532ms on this host; these include CLI/transport overhead and do not
  measure first-visible latency or establish a performance threshold.
- `VERIFIED`: only owned fixtures close normally, with unchanged saved bytes;
  seven original open documents and all six known saved originals remain exact.
  The previous Word PID exited externally before continuation. A new explicit
  instance restored the originals and was visible despite hidden startup; no
  hidden-session or clean-machine claim follows. Indexed COM inspection replaces
  intermittently incomplete pipeline enumeration. Retain the initial windowless
  bootstrap, premature Stop observation and duplicate-title constructor failures.
  The same Stop later settles without another Stop; the ambiguous constructor
  makes no target choice and closes normally before a uniquely named fixture.
- Evidence: `artifacts/word/preview8-installed-batches-isolated/qualification-result.json`,
  with the earlier failed packet in `preview8-installed-batches/`. Previous-share
  ownership cleanup and final local-only endpoint Stop are observed. The public
  endpoint is now stopped; future WebGPT testing needs an equivalent reconnect.
  Source, package and public assets are unchanged. Browser/rendering, normal
  latency, physical keyboard/custom typing, clean-machine and future Mac gaps stay open.

## Local resource diagnostics — source checkpoint, 2026-10-21

- `VERIFIED`: final source component run passes all twenty-one gates, including
  2,134 transport assertions and two effective diagnostic fault controls. Fifteen
  CLI checks pass, with the earlier unknown-command regression retained. Forty
  isolated native Word assertions verify local dispatch and preservation of text,
  active selection and foreground. Existing safety/transport controls remain in
  the component run. The aggregate `PARTIAL` leaves installed/public scope open.
- Counters cover parsed HTTP requests, HTTP refusals, RPC protocol errors,
  refusals carrying Origin, resource lists, read attempts/results/errors and
  completed successful resource response writes. Tests cover numeric-only data,
  unavailable/zero state, URI/Origin/protocol refusals, notifications, concurrent
  reads, stable local observation and lifecycle reset. A socket write is not
  proof of client receipt or widget bridge initialization.
- `VERIFIED`: native/disk audit after source tests preserves seven open documents
  and all six known saved originals. Evidence is private under
  `artifacts/word/resource-diagnostics/` and the preview.6 audit owner. The initial
  fault control failed compilation after removing its lock body; the corrected
  copied fault compiles and fails the intended observation predicate. The earlier
  six-counter prototype is not substituted for the final nine-counter result.
- At this source checkpoint, installed and Chrome counter baselines were `NOT RUN`;
  preview.6 lacks the command. The following preview.7 qualification covers the
  new installed bytes. Source adds no remote tool, permission, persisted logs or
  HTML/bridge change.

## Installed preview.7 CLI and real picker continuation — 2026-10-21

- `VERIFIED`: the distinct local candidate binds 72 source inputs and eight
  payloads. Exact-package COM 536 and CLI fifteen checks pass, including their
  effective fault controls. The 23,886,046-byte installer SHA-256 is
  `F5E85B7CCD2427AD1FF7ECD0B861ECD00E003F3CDB8AF6877E855AD737FB95CC`.
  The widget dependency manifest is refreshed for the version lock; packaged
  HTML remains the same 58,378 bytes and SHA-256 recorded below.
- `VERIFIED`: Word-open installer refusal preserves all eight installed payloads.
  Normal Stop reaches stopped/not busy/no pending cleanup before Word exits.
  Upgrade from preview.6 exits zero; payloads, manifest, version and both registry
  views match. All 140 profile files remain unchanged, with none added. Seven
  documents reopen with exact text, saved bytes, annotation counts and Saved state.
- `VERIFIED`: the exact installed CLI passes all forty-two isolation checks,
  including actual MCP wording writes, independent sessions, witness refusal,
  replay, revocation/recovery, local-only timer autosave and closed targets.
  There are zero unqualified foreground comparisons and no mouse/keyboard input.
  This candidate's pass does not rewrite preview.6's retained partial result.
  Nine separate installed diagnostics checks pass: inactive/zero/reset state,
  stable local observations, exact resource counts and refused-read separation.
- `VERIFIED`, resource observation: counters start at zero before equivalent
  connector creation. Discovery completes one successful resource result/write;
  the first cold conversation completes a second. Reload and visible document
  selection cause no further resource read. There are no resource/RPC/Origin
  errors in these snapshots; aggregate HTTP refusals remain separately counted.
  Completed writes do not establish host receipt or identify the cold failure.
- `FAIL`, first cold render: ChatGPT reports could-not-display, while its ordinary
  assistant response loads. `VERIFIED`, subsequent reload: the embedded picker
  visibly renders seven choices. Clicking the owned H document changes the UI to
  its selected title and automatically resumes the conversation. The assistant
  reports a complete read and preserves wording where a saved prompt overlaps
  Keep unchanged. Native CLI state confirms H shared and G unshared, with four
  marks/eight rules each. This is observed picker/continuation success after
  reload, not proof of a cold fix or captured bridge wire messages.
- `VERIFIED`: final audit preserves all seven open documents and all six known
  saved originals. Private evidence, screenshots, counter baselines and lifecycle
  receipts are in `artifacts/word/resource-diagnostics/preview7/`. Installed DLL
  SHA-256 is `BB5F03D1D241566786296C2BEE90D1BF3A3290CCCD1F1CE1675109E4C61B47F9`;
  CLI SHA-256 is `25F0EA9DBE035AD4F1377CFD6DA0F043E57665AB1E96E787E944AFDC417913F4`.
  Preview.7 is local only; public preview.5 assets are unchanged. Cold rendering,
  raw client wire, physical UI, first-visible timing, clean-machine and Mac
  acceptance remain open. Source checkpoint `838aeb6` passes hosted verification
  `39507792665`; verify the version/evidence checkpoint separately.

## Picker output declaration and installed preview.6 — 2026-10-20

- `VERIFIED`: missing-schema regression fails against the prior source. The
  corrected `test-transport.ps1 -MutationCheck` passes 2,117 HTTP assertions,
  fifteen JSON Schema positive/negative cases, and its existing safety mutation
  controls. The declaration closes successful list/choice fields and coded
  error envelopes, including nullable selection/title and opaque error details.
- `VERIFIED`: the actual installed preview.5 CLI listing matches that source
  declaration with seven open documents. An explicit owned H complete read and
  subsequent listing also match, retain the exact chosen identity and require no
  mouse/keyboard or wording write. Raw private results and retained red/green
  evidence are in `artifacts/word/picker-output-contract/`.
- `VERIFIED`: `modules/word/test.ps1` passes all twenty-one component gates,
  including exact built-assembly COM 536, transport/safety controls, autosave 50
  and native UI 1,419 checks with 21 renders. That source checkpoint's overall
  `PARTIAL` leaves installed/public-client scopes to the following candidate. A fresh native/disk audit
  passes for five open originals and all six known saved originals; no changes.
- Current [OpenAI plugin guidance](https://developers.openai.com/plugins/reference)
  asks for output schemas with structured results; the
  [MCP specification](https://modelcontextprotocol.io/specification/2025-06-18/server/tools)
  makes them optional. This is a documented host-contract correction, not proof
  of the blank picker's cause. HTML/bridge bytes, tool names, result fields,
  permissions and safety annotations are unchanged. Only `list_manuscripts`
  gains an output schema; the other three tools remain separate scope.
- `VERIFIED`: the separately identified preview.6 candidate binds 71 source
  inputs/eight payloads. Exact-package COM 536 and CLI fourteen checks, including
  effective fault controls, pass. The 23,885,251-byte installer SHA-256 is
  `6A2BD3C11AC2B3F2EF3DEE4A4B02C8D222632BBB8F3AD888D8F6BDE11451A17D`.
  Word-open refusal leaves all eight installed payloads unchanged. Real upgrade
  from preview.5 exits zero; all payloads, manifest, version and both registration
  views match. All 140 profile files remain byte exact across installation.
- `VERIFIED`: seven saved documents close/reopen under one new Word process;
  text, saved-file bytes, annotation counts and Saved state match. Five open
  originals and all six known saved originals remain unchanged after browser
  tests. The first preflight's nested JSON-array refusal is retained separately.
  The Stop call was asynchronous; closing before it settled left an encrypted
  ownership barrier. That barrier survived installation unchanged, and normal
  Stop after restart completes exact cleanup. No barrier was deleted manually.
- `PARTIAL`: `test-installed-cli.ps1 -IsolationCases` passes seventeen checks;
  Word holds the foreground, so positive writing/isolation cases remain `NOT RUN`.
  There are zero unqualified foreground comparisons. The earlier preview.5
  forty-two-check result does not certify these new installed bytes.
- `VERIFIED`: installed and public tool discovery exactly match the qualified
  four-tool source catalogue, including the new schema. A newly connected
  equivalent ChatGPT app retains Allow low-risk. Its real named-H fallback
  reports a complete read; native/disk audits confirm no document changes.
- `FAIL`: cold and reloaded real ChatGPT picker rendering still report
  could-not-display and response-load failure. The schema correction is not
  sufficient. Raw browser resource/wire/bridge initialization remains unproven.
  Console capture contains two Statsig cache warnings, without a proven cause.
- `VERIFIED`: the corrected public probe identifies itself as a non-browser
  CLI and receives three JSON HTTP 200 replies, the exact catalogue and the exact
  58,378-byte packaged HTML/MCP Apps MIME. The original default PowerShell probe
  instead received zrok warning HTML and is retained as a failed probe, not an
  MCP success. This does not establish what the ChatGPT widget loader received.
- Evidence: `artifacts/word/picker-output-contract/preview6/`; current process,
  endpoint and fixture details stay private there. Preview.6 is installed only;
  no release asset changed. Both rebuilt public page HTML files still exactly
  match deployed preview.5. Recheck this version checkpoint's hosted CI/Pages.

## Exact RGB recolor qualification — 2026-10-20

- `VERIFIED`: the existing policy recolors exact green `#00FF00` to red
  `#FF0000`, retaining ranges, text, unrelated marks/prompts and all prior rules.
  Fresh and existing neutral red have empty prompts; existing Reject red is
  reused unchanged. The expanded policy suite passes 201 assertions. A copied
  implicit-Reject mutation fails the intended new neutral-red assertion.
- `VERIFIED`: the new MCP-instruction regression fails against the prior
  guidance, then the updated guidance passes all 2,116 loopback transport
  assertions. Instructions and tool discovery explicitly permit neutral red
  without an editing-prompt question; no tool or permission schema changed.
- `VERIFIED`: `test-native-highlighting.ps1 -RecolorStress` passes 217 recolor
  assertions (221 including setup) in its own Word process. This includes exact
  green-to-neutral-red paint, empty prompt, replay, unchanged neighboring marks,
  text/fonts/rules/annotations and save/reopen, plus the existing overlap,
  refusal, Undo/Redo and inspection-budget scenarios. Its manifest has no input
  drift. Afterward, only the three test files' original LF endings were restored;
  `test-line-endings.json` retains before/after hashes for this formatting-only
  difference. Runtime source bytes were not changed by that restoration.
- Receipts and retained failures are private under
  `.for-ai-local/recolor-20261020/`. The first ad hoc policy compile used relative
  paths rejected by csc; its corrected absolute-path invocation passed.
  Installed/public-client behavior, physical UI and Writer parity are `NOT RUN`
  in this follow-up. No installer, release or site deployment was performed.
- `VERIFIED`: the final `modules/word/test.ps1` run passes all 21 component
  gates with the final source bytes: build, CLI, policy, concurrency, uncertain
  outcomes, recoloring, native review/mapping/input, COM/dispatch, transport
  negative controls, autosave, 1,419 native UI checks with 21 renders, and
  architecture. The driver's overall `PARTIAL` correctly leaves installed and
  public-client scopes unrun. Context structure/safety checks also pass, with
  existing oversized-document and dirty-tree warnings. Evidence is in
  `components-final/`; this is a local source checkpoint, not a deployment.

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

Installed continuation: `test-installed-cli.ps1 -IsolationCases
-AllowForegroundChanges` passes 36 checks on the same installed candidate
(`installed-isolation-active-desktop/`), with zero unqualified foreground checks.
This covers two actual loopback sessions, unnamed choice before text, retained
independent targets, cross-client witness refusal, same-ID recovery, direct edits
and replay in both documents, revocation/recovery, actual timer autosave and
closed-target refusal. The copied CLI session-routing fault fails the intended
unnamed-client isolation predicate after sixteen successful prefix checks.
Independent saved-DOCX audits pass twelve checks for exact Unicode, unchanged
neighbors, no revisions, one exact rationale/body anchor and green shading; a
copied rationale-content fault fails its exact predicate. The post-control audit
again preserves all six originals, including native state and file hashes.
See `isolation-qualification-result.json`, `session-identity-negative-control.json`
and `isolation-completed-originals-result.json` under the private evidence root.

Retained failures: the first extension assumed one marker after direct editing;
the actual state retains both original and Preserve bindings. The correction
compares every marker/rule before and after revocation. Later traces locate the
native position change at the harness's own Word Close call; CLI requests retain
the document/selection before closure and the native position after closure.
Transient closing windows and ChatGPT/Brave foreground changes are recorded.
Strict focus assertions remain the default. Explicit `-AllowForegroundChanges`
records any changed-foreground comparison as unqualified and makes the result
PARTIAL; it never relaxes wording, isolation or native edit guards. The final
36-check run used that diagnostic option but qualified all focus comparisons.
These are installed CLI/native and saved-file observations, not WebGPT, physical
UI, host reopen or Mac results. Local sharing is stopped after owned tests.

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
  pass. Hosted verification `39470432974` passes on CLI checkpoint `0aeb646`.
  All 71 committed installer inputs and the exact installed test driver match.
  Pages run `39470433033` fails the existing missing-font readiness race: the
  initial HTML already says unavailable before labels have measurement status.
  A test-only wait for label initialization preserves the fallback assertions;
  the corrected local rendered matrix passes all 313 checks. No site UI change.
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

At the preview.4 checkpoint it was a local installed candidate; the public
release/site were preview.3.
Subsequent WebGPT preparation exposes its local-only background autosave gap:
`local-autosave-installed-red-2/` fails the unchanged fifteen-second timer gate
before any MCP read. The prior 36 checks qualified saving after MCP selection.
The new source regression first fails background eligibility, then passes all
38 native assertions after successful local mutations enroll the target without
sharing MCP authority (`local-autosave-red/`, `local-autosave-green/`). These
source results do not qualify the old installed DLL. Initial preparation had a PowerShell `cli` alias collision;
the corrected preparation retained the actual timer failure. The first installed
hash probe used insufficient file sharing and failed before the intended gate;
its exact owned leftover fixture was closed with a verified path, and the driver
now tracks fixtures immediately and reads saved hashes with read/write sharing.
Concurrent recoloring source/test/document edits appeared during final review;
preserve them outside the CLI checkpoint and compare the candidate to the staged
Git inputs. Do not claim the whole working tree matches that installer. Physical
UI, full real WebGPT, clean-machine and Mac acceptance remain open.

### Preview.5 exact installed qualification

- `VERIFIED`: all 21 fresh component gates pass (`autosave-components/`),
  including 1,419 native UI checks/21 renders, transport 2,116 and autosave 50.
  The overall component `PARTIAL` retains its installed/public boundary. The
  source-native CLI suite passes 38 assertions with 47 exact bound inputs.
  Two packaged Markdown files changed after components; all executable inputs
  remain exact. Independent final-package checks cover the actual binaries.
- `VERIFIED`: preview.5 binds 71 source inputs and eight payloads. Installer is
  23,882,311 bytes, SHA-256
  `e54dc5e05246c292f8071d8ca3a533cdf94ffd5683f8f0286bd4d7e4ee8e0a80`;
  DLL `99d4c51fa39d539242907d59208b180201857e187fd8a3fa6655400d20e84550`;
  CLI `26b357fe2e0fd62eeec6df6539faf4182d81e34e4077d2dd1bb34a66203e9fc3`.
  Its exact DLL passes COM 536 assertions/effective dispatch control; exact CLI
  passes 14 checks/effective input control. See `preview5/`.
- `VERIFIED`: Word-open installation refuses with exit 2 and all installed
  payloads unchanged. Upgrade, uninstall and reinstall exit zero; all eight
  payloads/version/manifest and both registration views match. Uninstall checks
  ten registry-view/key combinations absent. All 133 profile files remain
  unchanged, none added (`preview5-lifecycle/`).
- `VERIFIED`: actual installed CLI passes 42 checks with zero unqualified focus
  comparisons, including inactive local-only timer saves before either MCP
  read, authority remaining unshared, two sessions, revocation, replay and
  closed-target refusal. Twelve independent saved-DOCX checks pass. A copied
  session-routing fault fails the intended unnamed-client isolation predicate
  after 22 successful prefix checks. See `preview5/installed/`,
  `saved-audit.json` and `session-identity-negative-control.json`.
- `VERIFIED`: five originals open at upgrade preflight retain native state;
  all six known saved originals retain exact bytes. The sixth was already
  closed; the cause is unestablished. The shutdown observer timed out, then the
  same Word process exited normally before installation. Reopen five and
  post-test preservation pass (`preview5-reopen/`,
  `preview5-originals-after-result.json`). No desktop input was used.
- `VERIFIED`, limited live scope: fresh owned G/H fixtures persist local marks
  before MCP access (`artifacts/word/webgpt-cli-qualification-3/`). The actual
  installed HTTPS MCP initialization returns HTTP 200 in 1,001ms. The equivalent
  ChatGPT app connects with four tools/unchanged low-risk permissions. A fresh
  named-document Instant chat edits G and leaves H unchanged. An independent
  unnamed chat reaches a blank picker cold and after reload, then a host
  could-not-display alert and response-load failure. The same chat's explicit
  H-name fallback completes H's edit. Twenty-two saved-DOCX checks pass for
  exact independent wording/Unicode, untouched intro/ending, green protection,
  one exact rationale/anchor and applied receipt each, with zero revisions.
  Five open/six saved originals remain unchanged. Client task times are nineteen
  seconds for G and eleven seconds for H's fallback; first-visible-edit latency
  is unmeasured. Browser control detaches while preparing an interleaved color
  request, before Send; both saved files retain their pre-attempt hashes.
  Interleaving/revocation and usable picker acceptance remain open. This host
  failure's cause is unestablished; it does not qualify a server-code diagnosis.
  Raw captures, capability URLs and original identities stay private.

- `VERIFIED`, preview shipment: public prerelease `v0.1.0-preview.5` targets
  `79168c326cd9926bdcaabb6c0dc104fba254da92`; hosted verification `39488257735`
  and Pages `39488257736` pass on that source checkpoint. All 71 committed
  source inputs and all three uploaded asset sizes/digests match. Anonymous
  HTTPS downloads return exact installer, build manifest and checksum bytes.
  The build-manifest SHA-256 is
  `9893e210a88515cd28f7fc984e9b3f088b4f6e7b6576d90dd4b924a9d908662a`.
  Earlier releases remain unchanged. See `preview5-release/` privately.
- `VERIFIED`, local site: release metadata uses the actual publication time,
  installer URL, SHA-256 and byte count. Updated release/limit copy retains
  existing layout and locked Pretext; the rendered matrix passes 313 checks.
  Minimum-width Description and 200%-text Download captures were inspected.
  Site checkpoint `3e9c7681c589103339baa3d8de297e32dace816d` passes Pages
  `39488997403`. Anonymous Description, Download, release metadata and manifest
  return HTTP 200; the three rendered documents match local tested bytes, and
  all twelve deployed source hashes match the published Git blobs. Hosted
  verification `39488997631` subsequently passes. Documentation checkpoint
  `b81eebfde6d46d9469ff20279e3b132378471939` passes verification `39489600126`.
  Recheck subsequent evidence-only publication state separately.

### Preview.5 real-client interleaving and reopen continuation

Private evidence: `artifacts/word/webgpt-cli-interleave/qualification-result.json`.
The overall packet is `PARTIAL` because the remaining host/platform scopes stay
open. Production sources, package, dependencies and persistent permissions are
unchanged; DLL/CLI hashes remain those above.

- `VERIFIED`: four actual Chrome/Instant requests interleave G/H neutral colors,
  including exact green `#00FF00` to existing neutral red `#FF0000` with empty
  prompt and unchanged rule definitions. Each operation passes eleven saved-
  DOCX assertions: exact wording/Unicode, unaffected marks/rules and rationale,
  native RGB, one receipt and the other conversation's exact saved bytes.
  Retained snapshots are bound to the executed audit hash in each result.
- `VERIFIED`, scoped refusal/recovery: after one-time host approval for the
  disposable protected-ending attempt, the client reports `INVALID_ARGUMENT`
  requiring native-identity preservation mapping. Both files remain byte exact;
  native observation still finds one rationale/no revisions/Saved=true each.
  A different valid color operation succeeds afterward. The raw tool envelope
  and original request identity were not captured; this does not qualify
  same-ID retry or real-client revocation.
- `VERIFIED`: the same installed Word process closes/reopens only owned G/H,
  hidden, without explicit Save or mouse/keyboard input. Eleven lifecycle checks,
  sixteen native/CLI semantic comparisons and two in-memory RGB checks pass.
  Review identities, text, comment content/author/date/scope, rules, marks,
  applied receipts and Saved state persist. The post-run audit preserves five
  open original documents and all six known saved original files.
- Retained harness failures: raw rationale XML equality fails because Word
  renumbers its export comment ID; the parsed difference is retained. Only that
  volatile ID is normalized, with full comment attributes/content and exact
  mapped body-anchor locations required. The missing-neutral-color observation
  control fails its intended predicate. Initial Close uses a value where this
  PowerShell COM binding requires a reference; it fails before either fixture
  closes. Both remain saved/open before the corrected fresh run. Neither failure
  is silently promoted to a pass or attributed to the add-in.
- `VERIFIED`, public resource boundary only: actual installed HTTPS initialize,
  resources/list and resources/read return HTTP 200 in 266/156/375ms. MIME is
  `text/html;profile=mcp-app`; returned HTML is exactly 58,378 bytes with SHA-256
  `416a21c068d4a179e112b90784c385032dafb1f98d2f11e505b5d7868b5119aa`.
  ChatGPT's Open app in tab fallback still reports could-not-display. Host
  receipt, iframe delivery and bridge/render initialization are unproven;
  resource HTTP success does not identify the cause or qualify a picker fix.

### Preview.5 real-client revocation, recovery and replay

Private evidence: `artifacts/word/webgpt-cli-revocation/qualification-result.json`.
The packet remains `PARTIAL` for the named host/platform and wire-capture gaps.
Production/package/dependency inputs and persistent permissions are unchanged.

- `VERIFIED`: the retained real H chat reads its chosen document completely
  after native reopen. The exact installed CLI then revokes H, changing shared
  true to false while preserving all rules, marks and pending proposals. The
  requested old-witness attempt is client-reported refused with `INVALID_ARGUMENT`
  and "Read this chosen document again before continuing its review." H stays
  unshared and both saved test files remain byte exact.
- `VERIFIED`, displayed-result/native correlation: after a fresh read the real
  client reports the original requested identity and replayed=false. Sixteen
  assertions correlate returned mark/color/replaced identities to exactly one
  new applied private receipt, its final native snapshot, exact saved RGB and
  unchanged text, other rules/marks, rationale and G bytes. A copied displayed
  marker-ID mismatch fails the exact native-correlation predicate.
- `VERIFIED`: the same chat's exact replay reports replayed=true and otherwise
  identical returned fields. Both DOCX files, the single owned encrypted ledger,
  document identity/revision and rules/marks/pending proposals remain exact;
  there is no duplicate persisted effect. Current native color/Saved checks and
  the post-run five-open/six-saved-original preservation audit pass.
- Retained audit failures distinguish two legitimate persistence stages.
  Document XML carries reserved highlight Undo staging and its precommit
  snapshot. The final applied outcome/snapshot reside in the CurrentUser-protected
  owned ledger, merged by `WordReviewStore.Load`. The first audit conflated the
  stores; a partial correction still compared a prior staged snapshot to its
  committed snapshot. Final checks require all other fields exact, authoritative
  committed outcomes, and snapshot hashes matching the captured native revisions.
  No production persistence or guard was changed to satisfy the checks.
- Scope limits: browser wire arguments/error envelopes were not captured.
  Request identity, error code/message and absent isError/uncertainty fields are
  displayed client reports. The host requested one-time approvals and flagged
  tool prose/test identifiers as suspicious; no persistent permission setting
  changed. These observations warrant description-clarity review, not a claim
  that the safety mechanism or native backend is defective.

Physical dialog, clean-machine, usable host picker and full client wire
qualification remain open. Mac implementation is deferred to the existing
future-agent handoff. The overall shipment goal stays active.

Raw browser captures, endpoint capabilities and document-preservation snapshots
remain private. Retain exploratory harness failures. Native source checks run
serially; do not change captured inputs while running. Native files intentionally
preserve CRLF under `.gitattributes`; review with Git's `cr-at-eol` interpretation.
