# Wordplai stress-test research

Research note, 2026-10-19. This is an engineering plan synthesized from
primary platform documentation and the repository's observed acceptance gaps.
It is not a Microsoft certification requirement, a published benchmark, or a
measurement of Wordplai's performance.

## Search and evidence method

The search began broadly across Windows UI automation, Office add-in testing,
installer validation, and browser test reliability, then narrowed to official
Microsoft and Playwright documentation on the exact proposed test layers.
Primary sources were opened where results exposed relevant detail. The
recommendations below distinguish platform capabilities from Wordplai-specific
acceptance evidence. No community/forum evidence was needed for this technical
question.

## Source ledger

| Claim used | Source and source type | Date/status | Confidence and qualification |
| --- | --- | --- | --- |
| Windows UI Automation exposes desktop controls to assistive technology and automated test clients; stable discovery depends on useful element properties and synchronization with dynamic UI. | [UI Automation](https://learn.microsoft.com/en-us/windows/win32/winauto/entry-uiauto-win32) and [Using UI Automation for Automated Testing](https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-usefortesting), Microsoft primary docs | Current pages; retrieved 2026-10-19 | High for framework capability. Does not prove Wordplai controls expose the needed tree or that pointer gestures occurred. |
| Accessibility qualification should combine automated checks in CI with focused keyboard and assistive-technology validation. | [Accessibility testing](https://learn.microsoft.com/en-us/windows/apps/design/accessibility/accessibility-testing), Microsoft primary docs | Current page; retrieved 2026-10-19 | High as Microsoft's guidance. The cited page addresses Windows apps generally; apply the principle to native WinForms controls. |
| Interactive UI automation can exercise desktop applications, but Microsoft does not recommend or support unattended, non-interactive Office automation due to instability and deadlock risk. | [Office automation considerations](https://learn.microsoft.com/en-us/office/client-developer/integration/considerations-unattended-automation-office-microsoft-365-for-unattended-rpa), Microsoft primary docs | Current page; retrieved 2026-10-19 | High for the stated boundary. A spawned, visible, PID-bound Word process is materially different from server/headless automation. |
| Office add-ins need testing on the Office versions/platforms users actually use; unit mocks cover logic but do not execute inside Word. | [Test Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/test-debug-office-add-ins) and [Unit testing in Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/unit-testing), Microsoft primary docs | Current pages; retrieved 2026-10-19 | High for Office Add-ins guidance; Wordplai is a native COM add-in, so use only the general principle and do not import web-add-in-specific tooling requirements. |
| Playwright traces preserve action, DOM snapshot, and network context useful for diagnosing browser-test failures. | [Trace Viewer](https://playwright.dev/docs/trace-viewer-intro), Playwright primary docs | Current page; retrieved 2026-10-19 | High for Playwright capability. A trace or successful synthetic bridge still cannot establish real ChatGPT host rendering or user-visible behavior. |
| Wordplai's acceptance needs distinct source, isolated Word, installed, real-client, and human-gesture observations. | Repository [HANDOFF.md](HANDOFF.md) and [VERIFICATION.md](VERIFICATION.md) | Local canonical project records, 2026-10-19 | High for the project contract; these are project acceptance rules, not external standards. |
| Stateful failures can depend on the ordering of inputs; pairwise or fixed-setting combinations do not cover all event histories. | NIST [Ordered t-way combinations](https://www.nist.gov/publications/ordered-t-way-combinations-testing-state-based-systems) (2022) and [Practical Combinatorial Testing](https://csrc.nist.gov/pubs/sp/800/142/final) (2010), primary testing-method sources | Published reports; checked 2026-10-19 | High for the method. Apply targeted ordered sequences alongside constrained parameter combinations; this does not guarantee defect coverage. |
| Word `Range.Text` returns plain, unformatted text and replaces the range when set; document stories are a separate collection; Word exposes a custom undo-record entry point. | Microsoft [Range.Text](https://learn.microsoft.com/en-us/office/vba/api/word.range.text), [Document.StoryRanges](https://learn.microsoft.com/en-us/office/vba/api/word.document.storyranges), [UndoRecord](https://learn.microsoft.com/en-us/office/vba/api/word.undorecord) | API pages last updated 2021–2022; checked 2026-10-19 | High for API semantics. Only actual Word fixtures can establish Wordplai's formatting, story, and Undo preservation. |
| Save As can change the saved document name or format and can overwrite a target path. | Microsoft [Document.SaveAs2](https://learn.microsoft.com/en-us/office/vba/api/word.saveas2) | API page last updated 2022; checked 2026-10-19 | High for API behavior. The expected retained-chat identity after Save As is a product policy that must be tested, not inferred from the API. |
| MCP cancellation can race an operation's completion and is not evidence that a native side effect was rolled back. | MCP [Cancellation](https://modelcontextprotocol.io/specification/2025-11-25/basic/utilities/cancellation), protocol specification dated 2025-11-25 | Versioned normative specification; checked 2026-10-19 | High for protocol semantics. Record the negotiated protocol version and reconcile the original request identity and native state. |
| An MCP App has distinct lifecycle stages beyond resource delivery; a resource response alone does not establish that the host received, initialized, rendered, or interacted with the UI. | MCP Apps [Overview](https://apps.extensions.modelcontextprotocol.io/api/documents/Overview.html), version 1.1.2 documentation | Versioned SDK documentation; checked 2026-10-19 | High for described lifecycle. Actual ChatGPT embedding and version behavior still require observation in that host. |
| UI Automation reports physical coordinate rectangles, and DPI scaling can complicate visual/automation correspondence. | Microsoft [Understanding Screen Scaling Issues](https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-screenscaling) | Current page; checked 2026-10-19 | High for UIA coordinate behavior. Font-size simulation is not evidence of a real monitor/DPI transition. |

## Three complementary testing approaches

1. **State-machine and transaction testing with retained fault controls.**
   Exercise authorization, target identity, current-context/revision checks,
   exact ranges, protection, replay, rollback, and uncertain outcomes against
   the existing component owners. Keep a failing copied control for each new
   critical guard. This gives fast, repeatable coverage of many interleavings,
   but cannot establish Word COM behavior, installed wiring, or rendered UI.
2. **Isolated interactive Word integration fixtures.**
   Spawn a visible Word process for each disposable synthetic fixture, bind
   automation to that exact PID, and capture native text, formatting, comments,
   revisions, saved bytes, and reopen state. Exercise two documents and two
   independent request identities, delayed responses, Undo/Redo, Save As,
   close/reopen, read-only and save failures. This observes the actual COM host
   and native document model, but still does not establish installed lifecycle,
   actual browser client behavior, or a physical human gesture.
3. **Installed end-to-end interaction and client observation.**
   After coordinated permission to install, run the packaged add-in through
   upgrade, clean install, use, uninstall, and supported/unsupported
   architecture checks. Use UI Automation for accessible names, keyboard
   traversal, and element invocation; separately have a human perform ribbon
   and text double-clicks and screen-reader checks. In actual fresh WebGPT
   chats, capture picker first load/reload, independent document selection,
   first visible edit and save latency. Keep browser traces/screenshots
   sanitized and private. This is the only layer that qualifies those named
   installed, physical, and live-client surfaces.

These approaches complement one another because each observes a different
boundary. A passing lower layer is not evidence that a higher layer passed.

For the first layer, combine constrained parameter coverage with explicit
event-order coverage. High-value histories include `read → local rule change →
delayed response → commit attempt`, `edit → response loss → exact replay`, and
`edit → user Undo → timer save/late response → reopen`. NIST's combinatorial
methods motivate this as a test-design technique; they are not a Wordplai
certification standard and do not establish that every relevant history has
been covered.

## Risk-ranked scenario shortlist

| Priority | Scenario | Required observation and negative control |
| --- | --- | --- |
| P0 | Two chats choose two open documents; windows activate in the opposite order | Exact target and changed passage/comment/green receipt per chat; other document remains byte/native-state exact. Repeat with duplicate titles and identical text. |
| P0 | Delayed response races with local typing, rule change, close/reopen, or Undo | Old response is refused before mutation; user edits and current rule remain intact. A valid fresh read can recover. Replay the same identity without duplicate edits/comments. |
| P0 | Protection and color identity under complex selections | Test neutral RGB, same-RGB distinct identities, exact recolor authorization, overlapping ranges, protected citations/links/fields, Unicode surrogate boundaries, and a late invalid target. Refusal is atomic; no rule gains authority. |
| P0 | Real host and picker cold/reload paths | In two fresh real chats, record resource preparation, host receipt, bridge readiness, first render, selected identity, first visible edit, and save separately. Preserve full actual errors privately. A server response alone is a failed qualification. |
| P0 | Installed foreground, keyboard, pointer, and accessibility behavior | Confirm exact refusal/no mutation while Word owns focus; move focus away and verify the intended path. Exercise keyboard-only selection, accessible names/focus, color-blind distinguishability, actual ribbon/text double-click, and overlap choice. |
| P1 | Installer lifecycle and environment | On an isolated environment, test clean install, upgrade, uninstall, interrupted install/recovery, Office x86/x64 applicability, unsupported OS/architecture refusal, and no active Word payload replacement. Compare installed payload hashes before/after. |
| P1 | Layout and recovery | Installed UI at supported DPI/scaling/font combinations, small working areas, and dialogs moved across monitors. No clipping or hidden Save; no-fit state remains readable, preserves draft, and recovers after resizing. |
| P1 | Document form and persistence | Synthetic legal review with tracked changes; grant tables/footnotes; academic citations; bilingual RTL/CJK/emoji; boilerplate and executive summary. Verify complete reads and saved/reopened native structures. Include duplicate titles, rename/move/Save As, read-only and failed save. |
| P1 | Transport and uncertainty | Drop the connection after a possible commit, late tool reply, cancellation while queued, or process restart. Retain original request identity and classify known refusal versus unknown outcome; never retry under a fresh identity to force success. |

## Additional high-risk interaction sequences

These rows refine the broader scenario matrix in [`HANDOFF.md`](HANDOFF.md)
with identity and ordering cases surfaced by the research pass. They are
proposals and have not been executed by this report.

| Priority | Trigger sequence | Required observation and forbidden result |
| --- | --- | --- |
| P0 | Two open documents have the same title and identical passage text; two chats choose different ones. | Picker identity and unique synthetic sentinel prove each write hit its chosen document; no quote-search or active-window retargeting. |
| P0 | Chat binds a document; Save As creates a new name; the old saved file is opened alongside it. | The documented target-identity policy is observed; only the intended live document changes. The old file is unchanged. |
| P0 | Chat chooses a document; it closes; a same-title copy opens before a delayed reply or old picker click. | Old authority does not transfer to the new instance; the original operation refuses or reconciles against its original identity. |
| P0 | User changes a local prompt rule after read and before delayed commit. | Old witness is revoked; no text/rule mutation occurs; a fresh complete read can recover. |
| P0 | User invokes Undo immediately before background autosave or a late response; then optionally Redo. | User's Undo remains authoritative, saved/reopened native text and marker/comment/receipt state agree, and late work does not reapply the edit. |
| P0 | A native modal dialog or busy Word operation is open while a remote command queues; connection then cancels or times out. | Bounded, stage-specific refusal or truthful unknown outcome; no false rollback claim, Word hang, or cross-document callback. |
| P1 | User enters text through paste, IME or dictation during a delayed operation; keyboard-only user opens and saves a rule dialog. | Input is preserved; focus guard is exact and bounded; full keyboard order and visible focus work without pointer input. UIA presence alone does not prove task completion. |
| P1 | Dialog moves across monitors with different scale factors; small working area and large readable text force no-fit. | No clipped/hidden actions; draft survives; no-fit recovery restores editing. Actual monitor transition is separately observed from font simulation. |
| P1 | Synthetic article with live citation fields, contract with tracked edits, grant table/footnote, Arabic/CJK/emoji, headers, footnotes, comments and page breaks. | Complete-read limits are explicit; supported edits preserve native structures and protected units; unsupported writes refuse atomically. |
| P1 | Clean-machine install, upgrade, uninstall, then supported/unsupported Office bitness and OS/architecture checks. | Installed DLL and payload hashes match the candidate; refusal happens before partial mutation; Word is never terminated or user data changed. |

## Measurement policy

For latency, define timestamps before running the scenario: user invocation,
dispatch admitted, first visible changed text, rationale/green visible, tool
return, and native save/reopen. Report each interval and environment. The
project's target is immediate passage-level editing; no external standard or
benchmark threshold was found in the reviewed sources. Establish a proposed
threshold explicitly as a product policy, then report observed distributions
separately. Do not label request-to-return harness timings as first-visible
real-client latency.

One external research draft proposes ≤200 ms p95 for local picker feedback and
≤2 s p95 from admitted native edit to visible wording/rationale/protection, and
suggests at least 100 independent trials for an initial p50/p95
characterization. These are experiment hypotheses only: they are not adopted
product SLAs, standards, acceptance gates, or measured Wordplai results. First
baseline on named synthetic fixtures; record cold/warm host, model/build,
document size, and separate native dispatch, first-visible, tool-return, and
save/reopen intervals before deciding whether to adopt a target.

For every scenario record `VERIFIED`, `PARTIAL`, `BLOCKED`, or `NOT RUN`, exact
surface, fixture class, source/package identity, result, negative-control
outcome, and evidence location. Counts overlap; report scenario rows
individually rather than summing assertions into a release score.

## Limits and next dependency

The official sources establish testing capabilities and cautions, not a
Wordplai pass result. The existing handoff says the browser connector returns
`FORBIDDEN`, browser profile binding times out, and install requires the user's
Word documents to be saved and closed first. Until access is restored and the
Word installation window is coordinated, those live checks remain blocked or
not run. No installer was run and no Word process or user document was touched
during this research.
