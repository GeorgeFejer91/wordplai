# Agent workflow, context maintenance, and Git policy

## Working loop

1. Read `AGENTS.md`, then the router, shared baseline in `SKILLS.md`, durable
   goal in `PROJECT.md`, and active acceptance checklist in `HANDOFF.md`.
   Reuse existing GSD state when present; keep one task-state owner.
2. Inspect Git status, branch, remotes, existing worktrees and task-owned files.
3. Preserve user work and distinguish pre-existing failures from regressions.
4. Apply YAGNI: check the existing owner before adding anything.
5. Implement one bounded outcome and verify it proportionately.
6. Review the diff for accidental scope, secrets, generated output, and stale
   context.
7. Update the existing task owner with evidence, blockers and the next concrete
   action. Report observed branch/commit/publication state; recheck it next time.
   Keep overall acceptance open until its own gates pass. Publish only when
   the rules below apply.

## Control-plane update triggers

Update the narrowest canonical file in the same change when any durable fact
changes:

- project goal, non-goal, user workflow, or supported surface;
- module ownership, architecture boundary, external contract, or schema;
- required skill route or verified fallback;
- acceptance criterion, validation command, evidence limit, or release gate;
- security, privacy, safety, licensing, or publication boundary;
- durable decision or repeatedly rediscovered failure mode.

Do not store chat transcripts, daily narration, raw logs, private machine paths,
speculative ideas, secrets, participant data, or facts obvious from source and
tests. Keep short last-observed Git state only in the existing task handoff
owner when needed for continuity; fresh Git state remains authoritative.

## Git checkpoints and publication

- Inspect status and diff before and after edits.
- Stage explicit paths; never absorb unrelated work to make the tree clean.
- Keep one coherent concern per commit and checkpoint meaningful milestones.
- Commit coherent, validated work. Push only when requested or covered by
  continuing explicit authorization, under the repository's branch/PR policy.
  The shared baseline supersedes the former automatic-push default.
- Never force-push, rewrite shared history, bypass branch protection, bypass
  secret scanning, or publish known-broken/private material.
- The user explicitly authorized this standalone preview with remaining
  acceptance work handed off. Label that release as a preview, expose its
  concrete limitations, and retain failed tests as failures. This does not
  authorize calling the product production-ready or publishing private evidence.
- If credentials, network, CI, conflicts, ownership, or validation block
  publication, report the exact blocker and preserve the ready local state.
- After pushing, verify the remote SHA and applicable CI/deployment. A push by
  itself is not deployment evidence.

## Self-review

Run `for-ai/scripts/check-context.ps1`. Then review control-plane changes for
contradictions, duplication, stale facts, oversized always-read material, and
rules that lack a current consumer. Consolidate or delete rather than append
forever.

## Itemized handoff

- Outcome and affected product surfaces.
- Intended files changed; unrelated work left untouched.
- Verification run, result, and evidence limitations.
- `for-ai/` update: yes/no and why.
- Commit SHA, branch, remote sync, CI/deployment result.
- Remaining blocker or next smallest product decision.
