# Current CLI implementation scope

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
