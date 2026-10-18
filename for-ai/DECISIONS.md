# Decision log

Record only durable choices whose rationale future agents would otherwise have
to rediscover. Source and tests remain the authority for implementation facts.

## D-0001 — Minimal control plane before product scaffolding

- Date: 2026-10-18
- Status: Accepted
- Context: The project needs a predictable AI entrypoint without choosing an
  application stack or inventing product structure prematurely.
- Decision: Use a short root `AGENTS.md` that routes to a lowercase `for-ai/`
  control plane. Keep product output outside that folder and add project files,
  skills, and protocols only for current requirements.
- Consequences: New agents get a reliable map and readiness gates. The first
  product task must still choose the smallest suitable output structure.
- Supersedes: None

## D-0002 — Public standalone preview with an explicit acceptance handoff

- Date: 2026-10-18
- Status: Accepted
- Decision: The latest user direction authorizes a public standalone repository,
  installer and two-page site now, with unfinished functions/stress tests handed
  off. This supersedes waiting for all inherited acceptance gaps before extraction.
- Consequences: Publish a clearly labelled preview. Preserve the foreground guard
  and report its refusal, real-client/gesture gaps and unverified installer scope.
  Copy no user documents, private connection data, captures or generated diagnostics.

## D-0003 — Preserve the native implementation and upgrade identities

- Date: 2026-10-18
- Status: Accepted
- Decision: Keep C#/COM/WinForms and the legacy assembly, ProgID, CLSID and profile
  identities. Preserve qualified native/resource bytes through Git attributes.
- Consequences: No UnderTree installation or source checkout is needed to build
  or run this add-in. It upgrades the same Word add-in identity rather than loading
  a second conflicting handler. Native behavior changes require their own tests.

## D-0004 — Stable download page, versioned preview assets

- Date: 2026-10-18
- Status: Accepted
- Decision: Keep `/wordplai/download/` permanent and publish installers as GitHub
  release assets. The page uses reviewed version/hash metadata, not the latest
  release alias, which excludes prereleases.
- Consequences: New versions update the page metadata without changing its URL.
  Keep third-party notices and the inherited UNLICENSED source status; publication
  is not an invented new open-source license grant.

## Record format

For later decisions, add one compact entry with:

- identifier and title;
- date and status (`Proposed`, `Accepted`, `Superseded`, or `Rejected`);
- context;
- decision;
- consequences;
- supersession link when applicable.

Do not rewrite accepted history to hide a changed direction. Add a superseding
decision and link both entries.
