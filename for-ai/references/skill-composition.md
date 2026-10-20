# Combining the project skills

Apply this housekeeping baseline to coding and software development in every
project. Select task-relevant workflows before acting and load only the references needed for that task. Preserve the project's
stack, scientific rules, security boundaries, and existing source of truth.

## Responsibilities

| Layer | Owner | Contribution |
| --- | --- | --- |
| Project context and routing | `for-ai` | Entry points, durable contracts, skill selection, Git policy, and evidence requirements |
| Command-output efficiency | `rtk` | Compress supported shell output; keep exact reads and machine input raw |
| Project lifecycle | GSD Core | Requirements, milestones, phases, execution state, and acceptance evidence |
| Engineering practices | Superpowers | Task-specific planning detail, debugging, TDD, review, worktrees, and completion checks |
| Small implementation scope | `ponytail` | Reuse existing owners and implement the smallest complete change |
| Product architecture and interface | Relevant Tauri and Pretext skills | Rust authority, native boundaries, HTML geometry, and rendered verification |

RTK reduces command output. GSD reduces repeated project discovery and organizes
work across contexts. Superpowers provides focused procedures inside that work.
Do not expect RTK to reduce the cost of loading unnecessary plans or skill files.

## Every-project entry and task size

1. Read project instructions and skill routes. Inspect status, branch, upstream,
   and existing worktrees. Preserve pre-existing changes.
2. Apply `rtk` for suitable readable shell output. Use raw reads for complete
   selected instructions, editing, parser input, and exact verification.
3. Read `superpowers-using-superpowers` to select matching engineering practices.
   Resolve upstream `superpowers:<name>` references to the installed
   `superpowers-<name>` skills. Never overwrite similarly named local skills.
4. Select the GSD path. For an existing `.planning/`, use its current state and
   `gsd-progress`. For a substantial new project use `gsd-new-project`; for an
   existing codebase use `gsd-onboard` when onboarding is part of the request.
   A small correction can apply GSD's explicit outcome, bounded scope, and
   evidence discipline inline without creating a milestone or extra agents.
5. Choose the product skills below. Use one plan, one task-state owner, and one
   integration decision for the current change.

Apply the baseline to software work inside research and writing repositories
too. Preserve their domain-specific evidence and publication rules.

## Goals and continuity across sessions

At session start, read the project's durable goal and acceptance criteria,
then resume its active checklist. Use `for-ai/PROJECT.md` for the product
contract unless an existing document already owns it. If GSD's
`.planning/PROJECT.md` owns the goal, link to that document from the For-AI
contract; do not maintain two competing goal descriptions. GSD execution
state lives in `.planning/STATE.md`, `ROADMAP.md`, and the active phase plan.
Follow the repository's existing mapping when it uses different names. Use
`gsd-resume-work` for a GSD resume and `gsd-progress` to select the next step.

If no GSD execution state exists, use the existing project's task checklist.
If there is no checklist, place a small active checklist in `PROJECT.md` until
the work needs phase management. Give each item an observable result and an
acceptance check. Avoid a second checklist that repeats GSD or an existing
specification tool's task state. Record task ownership when agents work in
parallel and keep dependent work ordered.

During work, mark an item complete only after its named acceptance evidence
exists. Link to that evidence and preserve failed or blocked outcomes. A local
task's completion does not establish the whole project goal, deployment,
physical-device behavior, or user acceptance. Keep the overarching goal open
until its own acceptance criteria are met.

Before ending a session or handing off, update the active state with:

- the overall goal and current milestone or task identifier;
- completed checklist items and the evidence that supports them;
- the next concrete unfinished action and its acceptance check;
- blockers, unresolved decisions, and relevant continuing authorization;
- the canonical repository, worktree/branch, commit, dirty paths, and remote
  publication state as last observed.

Use `gsd-pause-work` when that workflow owns the state. Keep the handoff concise
and route long evidence to files rather than copying logs. The next session
must check fresh Git state before acting; saved branch or runtime observations
can become stale. Reuse this state instead of relying on chat memory.

## Planning and state overlap

GSD owns active milestone and phase state under `.planning/` when that workflow
is in use. `for-ai/PROJECT.md` and `DECISIONS.md` hold durable project contracts
and decisions. Keep their responsibilities distinct. Link phase conclusions
into durable context only when a contract actually changes.

Use `gsd-discuss-phase`, `gsd-plan-phase`, and `gsd-execute-phase` for phase work.
Use `gsd-review` and the relevant verification workflow to check the result.
The standard install exposes `gsd-help` and `gsd-surface` for other capabilities;
inspect their installed guidance before selecting an unfamiliar command.

Superpowers brainstorming can resolve an unclear decision before a GSD phase.
Its planning technique can improve task boundaries inside the GSD plan. Set
the plan location to the active GSD phase instead of creating a duplicate
`docs/superpowers/plans/` plan. Its execution ledger must reuse GSD's active
task state. If an existing SDD/spec workflow already owns the change, preserve
that authority and link the GSD coordination context; do not migrate formats
or create competing task lists without a user request.

Respect design decisions and execution authorization already supplied by the
user. A toolkit's default approval menu does not cancel that authorization.
Use actual harness tools and model availability. Do not enable features,
change model configuration, or fabricate missing agent roles from a template.
Do not launch a second agent harness as a workaround for unavailable dispatch.

## Superpowers routes and overlap

The local installation prefixes these identifiers with `superpowers-`.

| Installed skill suffix | Apply when | Combine with |
| --- | --- | --- |
| `using-superpowers` | Selecting the workflow at task entry | This routing document and the project's instructions |
| `brainstorming` | Product intent or a design decision is unclear | GSD discussion; retain one agreed decision |
| `writing-plans` | Requirements need implementation tasks | GSD's active plan location and task state |
| `executing-plans` | Implementing an agreed plan in the current session | GSD phase plan; reuse its state |
| `subagent-driven-development` | Authorized task-level delegation has a clear benefit | One GSD dispatcher and disjoint ownership; avoid nested orchestration |
| `dispatching-parallel-agents` | Independent tasks justify authorized parallel work | Existing worktrees and the selected lifecycle owner |
| `using-git-worktrees` | Substantial independent work needs isolation | The Git policy below and native Codex worktree ownership |
| `systematic-debugging` | Reproducing and explaining a defect | Local `debugging`; share one hypothesis and evidence log |
| `test-driven-development` | Implementing behavioral code changes | Local `writing-tests`; use one regression cycle, preserve existing gates |
| `requesting-code-review` | A change needs review before integration | Local `code-review` and GSD review; avoid duplicate review reports |
| `receiving-code-review` | Assessing incoming review comments | Local `receiving-feedback`; verify the claim before changing code |
| `verification-before-completion` | Making a completion claim | GSD acceptance evidence and project verification gates |
| `finishing-a-development-branch` | Integrating completed work | Existing user instructions about push, PR, merge, and cleanup |
| `writing-skills` | Skill behavior needs improvement | `skill-creator` owns package format; `optimize-skills` owns concise routing |
| `diagnosing-superpowers` | Investigating workflow cost or compliance | Inspect the actual run; apply only supported conclusions |

Read the selected skill completely. An available toolkit is not a requirement
to execute every workflow. For a task with an existing outcome and sufficient
evidence, avoid a second planning, review, or verification ceremony. Repeat a
check when a changed input or unresolved risk gives it a purpose.

## Rust, Tauri, HTML, and Pretext

| Product surface | Primary skill | Additional guidance |
| --- | --- | --- |
| Existing Rust/Tauri app or native IPC work | `tauri-rust-developer` | Load only affected architecture, Rust quality, security, and verification references |
| New or broadly migrated app with CLI and browser companion | `tauri-remote-app-builder` | Own the integrated app contract; preserve the user's chosen scope |
| Adding or qualifying a remote browser/phone companion | `tauri-browser-remote-control` | Own authentication, authority, state, transport, and real remote-path evidence |
| Text-bearing HTML/CSS/browser or Tauri WebView UI | `uncodixfy-pretext` plus `ponytail` | Required Pretext measurement for touched bounded text and rendered DOM checks |
| Bounded resizable app control panel | `uncodixfy-pretext`, accordion stretch mode | Read its `references/accordion-stretch.md`; declare supported minimum dimensions |

Prefer accordion stretch when the requested panel must expand and compress in
both dimensions without panel scrolling. Preserve header/action anchors, group
order, readable minimum type, and full-text access. Measure the usable inner
boxes with Pretext and verify the actual DOM across width/height combinations,
zoom, long strings, and supported WebViews. Clipping or hiding scrollbar chrome
does not satisfy that contract. Ordinary document pages keep document flow.

`uncodixfy-pretext` owns the adapted Uncodixfy aesthetics and geometry for these
interfaces. Consult `uncodixfy` as an additional visual reference when useful;
do not apply conflicting palette or font examples over the product's design.
The app-builder's Pretext guidance and this UI skill share one typography
contract and one verification matrix. Do not implement separate measurement
wrappers or load every Tauri reference to cover the same surface twice.

## Agents Inc and application state

Read [agents-inc-composition.md](agents-inc-composition.md) for the selected
companions and their overlap. Use `mobx-integration` when evaluating MobX or
changing a MobX state boundary, `agents-inc-accessibility` for touched HTML
interaction semantics, and `agents-inc-web-performance` for measured browser
or WebView costs. Preserve the existing stack and state owners. These routes
do not make MobX an application dependency in every project.

## Git and worktree policy

Git is the version and rollback record. Use one canonical repository and its
registered worktrees. Never create `_backup`, `_old`, timestamped source copies,
or duplicate project folders to preserve code versions.

For substantial independent work, prefer a dedicated branch and worktree when
isolation is useful. This is a standing worktree preference; do not ask for it
again. Inspect current isolation first, distinguish submodules from linked
worktrees, and reuse Codex-managed worktrees. Prefer native management when
available. Otherwise use Git worktrees and verify a project-local worktree
directory is ignored before creating one. A small in-place instruction change
does not require a new worktree.

Use the project's locked dependency setup and relevant baseline checks. Generic
toolkit setup examples do not require reinstalling dependencies or rebuilding
every project for a documentation change.

Make small coherent commits after proportionate verification. Keep `main`
stable and respect branch/PR rules. Push only when the user requested it or a
still-applicable explicit authorization covers it. Do not force-push, reset
destructively, discard uncommitted work, or remove worktrees without explicit
authorization. A directory name alone does not prove cleanup ownership.
Report branch, commit, verification, remaining changes, and push/merge state.

## Installation and upgrades

Read [toolkit-provenance.json](toolkit-provenance.json) for reviewed source
revisions, installed names, and adaptations. Keep upstream licenses. Review
future instructions, helpers, dependencies, and permission changes before an
upgrade. Missing tools require an honest fallback, not a claim that they ran.
Do not enable optional API proxies, telemetry, visual-companion servers, update
hooks, account operations, or publication merely to activate project guidance.
