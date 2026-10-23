# Optional compact Secret Tunnel companion

## Requested outcome

Offer three Windows x64 preview installers from Wordplai's live GitHub download
page: Wordplai alone, a compact Secret Tunnel alone, and both applications in one
installer. The companion is optional; Wordplai remains independently usable.

The user requested a parsimonious main page containing zrok setup with an account
registration link, GitHub connection, and read/edit rights. A second Folders tab
supports selecting multiple folders. Reuse the existing Secret Tunnel native
authorities and runtime; do not create a second file/GitHub implementation.

## Companion surface

- Product name: Secret Tunnel Lite (implementation choice, not a user-supplied name).
- Connection tab: status, copyable MCP URL, Start/Stop, zrok account link and masked
  enrollment token, GitHub device-flow connection, and Read / Read and edit /
  GitHub access choices. Explain the permissions before enabling them.
- Folders tab: choose the primary GitHub workspace, add/remove additional approved
  roots, show full paths with the existing measured paging behavior.
- Exactly two tabs. No network, storage, profiles, skills, custom system-prompt,
  Codex, voice or alternate-provider pages. Use the existing Pretext/layout owners,
  semantic controls, visible focus and status announcements. No scrollbars.
- Reuse existing Rust validation, credential protection, operation guards,
  lifecycle, autoscope and GitHub coordinator. Hidden controls alone do not
  establish an authority boundary: omit optional native brokers and remote tools.
- Use an independent application identity/profile and an independent zrok
  enrollment. Ship no tokens, private profiles or inherited folder permissions.
  Do not copy another installed environment's identity into a fresh Lite profile.
- Keep the full Secret Tunnel build working. Implement Lite as a build flavor in
  the canonical SecretTunnel-v2 repository, preserving shared backend ownership.

## Connection choice

The preview keeps separate Wordplai and Lite MCP endpoints. The optional question
about a shared URL remains unanswered; independent endpoints satisfy the explicit
three-package request with no forwarding authority or routing layer. The user was
told this assumption before publication. Copying an account token does not create
a shared URL. A future shared endpoint requires its own scope/revocation design.

## Packaging and publication

- Keep the exact published Wordplai preview.8 installer and its verified SHA-256
  as the Word-only component; current preview.9 source has separate release gaps.
- Build an unsigned Lite preview with its exact source revision and a compact
  runtime allowlist: native executable, Node, zrok, MCP server and required notices
  and readiness probe. Exclude optional runtime downloads and the offline WebView2
  installer; use Tauri's normal WebView2 bootstrapper for missing installations.
- The combined per-user installer embeds the exact two standalone installers.
  Preflight open Word, run each component once, report a failed component and a
  nonzero exit code, and preserve installed applications/user data on partial
  failure. Each component retains its own uninstaller. Never claim an atomic
  rollback that did not happen.
- Publish a clearly labelled preview release with all three named installers,
  exact checksums, provenance, requirements and observed qualification limits.
  The live Wordplai Download page presents the three choices and their distinct
  versions, sizes and hashes. Preserve the original release/download.
- The user explicitly requested live GitHub publication of these installers.
  This authorizes the related commits, pushes, release assets and Pages update;
  it does not authorize signing with an unspecified identity or unrelated releases.

## Acceptance

1. First-run, configured, connecting/error, GitHub-device and many-folder states
   fit both tabs at normal/minimum widths, 320 CSS px reflow, enlarged text and
   text-spacing overrides; all values/actions remain accessible through paging.
2. Test the real shared permission and folder logic, plus Lite profile separation
   and absence of optional remote brokers/tools. Test cancellation and narrowing
   access with active clients. Separate synthetic UI checks from native behavior.
3. Build and inspect the exact package allowlist; no private profile, account
   credentials or unrequested large runtime is included.
4. Qualify standalone/combined install and uninstall using isolated profiles and
   explicit application paths. Preserve the user's Word documents and live
   Secret Tunnel/Wordplai profiles. Record unavailable clean-machine evidence.
5. Download all three public assets and compare bytes/checksums, inspect the live
   three-choice page, and verify applicable hosted checks/deployment.

Do not call this a stable release or claim a real ChatGPT conversation succeeded
from a loopback/HTTPS protocol probe alone.
