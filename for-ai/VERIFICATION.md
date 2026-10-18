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
  native UI checks and 21 renders. Hosted qualification remains separate.
- Browser control timed out twice, and the saved Wordplai connection returned
  `FORBIDDEN`. Real WebGPT behavior and timing remain blocked pending a working
  browser/connection; no retired connector or simulated client substitutes for
  that acceptance evidence. No capability URLs are included in this report.
