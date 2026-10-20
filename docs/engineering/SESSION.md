# Current CLI implementation scope

## Highlight recolor follow-up

Qualify exact green `#00FF00` to red `#FF0000` recoloring through the existing
policy/native owners. Preserve text, exact ranges, other markers and all existing
rule definitions. Red may be neutral with an empty prompt; hue never assigns an
editing function. Clarify MCP discovery guidance and extend existing policy,
transport and native recolor checks. No new tools, rule migration, UI, installer
or publication changes. The CLI changes listed below predate this task and must
be preserved. Use owned isolated native fixtures only; installed/live-client and
Writer parity evidence remain separate. Initial branch is main at `3f6e18a`,
with the pre-existing CLI/product/documentation changes uncommitted. No new
failure has yet been observed for the requested color behavior.

The user requests CLI functions to automate add-in testing without desktop input.
Start from clean main `3f6e18ad6096011992091c592d01ec23c7348646`.
Reuse the native local COM authority and existing test drivers. Add explicit
background-document targeting and range marking for local automation, without
adding remote MCP tools, weakening complete-read/protection/replay guards, or
using mouse/keyboard/foreground injection in the CLI.

Allowed files: WordEditor/WordLocalReview/ComAddin and a focused automation partial;
modules/word/cli, matching native/client tests and drivers, build/packaging,
module/root documentation, and the existing for-ai project/handoff/verification
owners. Preserve unrelated work and all user documents. Use only owned synthetic
fixtures for write tests. Retain preview.3 release assets unchanged; any new
packaged candidate needs a new version and independent identity/installed checks.

Validate missing/ambiguous/stale targets, exact UTF-16 ranges/quotes, inactive
document isolation, selection/focus preservation, rule/marker persistence, JSON
and exit-status behavior, and real installed CLI commands. Local commands use
local user authority; remote AI changes continue through the existing MCP tools.
CLI results do not certify pixels, physical gestures or real WebGPT rendering.
Mac implementation stays deferred to its existing future-agent handoff.

CI follow-up permits a test-only fix in `site/test.mjs`: wait for missing-font
fallback labels to initialize before checking geometry. Preserve the actual
fallback assertions and rendering; retain the original hosted failure.

Installed CLI continuation adds opt-in two-session isolation, cross-client
witness refusal, local revocation, closed-target refusal and timer-saved DOCX
checks to the existing driver. Change only that driver and narrow CLI test
documentation/evidence. Use the exact installed candidate, owned hidden fixtures
and explicit existing Word PID; preserve concurrent recolor work and originals.
