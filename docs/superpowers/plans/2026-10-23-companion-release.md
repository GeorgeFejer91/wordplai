# Companion release implementation plan

> **For agentic workers:** Use superpowers-executing-plans to implement this plan
> task by task in this session. Keep one integration owner.

**Goal:** Publish Wordplai-only, compact Secret Tunnel-only and combined Windows
preview installers with a verified live three-choice download page.

**Architecture:** Add a Lite build flavor to the canonical SecretTunnel-v2 v5
backend and a separate two-tab frontend. Wordplai owns the combined NSIS package
and download catalogue. Component installers and profiles remain independent.

**Tech stack:** Existing Rust/Tauri 2, TypeScript/Vite, Pretext, Node MCP runtime,
PowerShell and NSIS. No framework migration or new app runtime dependency.

**Spec:** ../specs/2026-10-23-optional-secret-tunnel-design.md

## Global constraints

- Windows x64 preview; unsigned; no inherited credentials or folder grants.
- Preserve existing applications, profiles, document contents and full builds.
- Exactly two compact tabs; measured readable content and paging; no scrollbars.
- Wordplai-only component is the unchanged published preview.8 installer.
- Resolve the pending shared/separate MCP endpoint choice before forwarding work.

## Review focus

- Narrowing permissions must withdraw active write access before reporting success.
- Long folder paths, 200% text and connection failures must preserve action access.
- New Lite profiles must not adopt another app's environment or existing grants.
- Bundle failures and open Word must never produce a false successful install.
- Public metadata and downloads must refer to the exact qualified bytes.

## Task 1: Lite flavor and two-tab UI

Repository: SecretTunnel-v2, isolated codex/lite-companion worktree.

Files: app/src/lite.ts, app/src/lite.css, app/lite/index.html,
app/vite.lite.config.ts, app/src-tauri/tauri.lite.conf.json,
app/src-tauri/{Cargo.toml,src/lib.rs,src/settings.rs,src/process.rs},
app/scripts/{prepare-lite-runtime.mjs,smoke-lite.mjs}, app/package.json.

- [x] Add a failing rendered test for the two-tab workflow, native command calls,
  read/write narrowing, many-folder paging, long errors and measured text.
- [x] Add focused Rust assertions for the Lite profile and optional-service gates.
- [x] Implement the smallest build flavor and UI using existing native commands,
  existing ui-layout/ui-pages, native GitHub device flow and permission operations.
- [x] Prepare only the pinned Node/zrok/MCP runtime and required notices/probe.
- [ ] Run frontend build, focused behavior/geometry checks, baseline UI checks,
  Rust tests, and package/resource inspection. Retain initial failures.
- [ ] Review/commit the coherent source changes and capture exact source identity.

## Task 2: Endpoint behavior and installed Lite qualification

- [x] Incorporate the user's endpoint choice into the spec. If shared, define
  forwarding/session/resource/revocation boundaries before implementation and add
  actual two-client tests. Do not proxy arbitrary upstream URLs.
- [ ] Build the exact Lite NSIS preview with an independent profile/application ID.
- [ ] Inspect launch/status on an isolated native profile and verify the exact
  installed binary/resources. Exercise disposable selected folders and real
  permission refusal; do not grant production access for a UI fixture.
- [ ] Qualify uninstall/profile preservation and document native/clean-host limits.

## Task 3: Combined package and three-choice catalogue

Repository: wordplai, isolated codex/companion-release worktree.

Files: modules/companion/{build.ps1,installer.nsi,README.md}, site/{build.mjs,
release.mjs,release.json,test.mjs,style.css}, README.md and for-ai/HANDOFF.md.

- [x] Add a failing check for three distinct installer choices and exact release
  metadata; add a package check for component errors and open-Word preflight.
- [x] Embed hash-verified standalone Wordplai/Lite components in the NSIS bundle.
- [ ] Build and qualify the bundle, partial failure, preservation and independent
  uninstall semantics. Keep private evidence under ignored artifacts/.
- [x] Update the existing download page with three accessible, measured choices.
- [ ] Run package checks, rendered site checks and source/context checks.
- [ ] Review/commit source and manifests without private evidence or binaries.

## Task 4: Publish verified previews

- [ ] Push related branches/source references under the user's release authority;
  inspect hosted checks and fix related failures before promotion.
- [ ] Create draft preview releases containing exact tested installers, SHA-256
  checksums and source/component manifests; review requirements/known limits.
- [ ] Publish release assets, update exact metadata and deploy the live Pages site.
- [ ] Download each public artifact, compare hashes, inspect the deployed choices,
  and record actual Git/release/deployment state in the active handoff.

The existing HANDOFF.md owns overall acceptance and next-action state. This plan
owns the implementation checklist; do not create another competing state tree.
