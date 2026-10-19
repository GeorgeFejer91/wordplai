# Cross-platform feature parity contract

## Purpose and authority

LibrAI is the LibreOffice Writer counterpart to [WordPlai](https://github.com/GeorgeFejer91/wordplai). These repositories have one shared product promise: every user-facing writing-assistant function offered by either sibling must be available in the other with equivalent user outcomes. This requirement applies in both directions as either product grows.

The host applications may require different implementation protocols, user interfaces, storage formats, APIs, packaging, and validation protocols. Do not force shared code or identical interaction when a host-native implementation is safer or more reliable. Preserve equivalent behavior, user control, protection, and evidence quality.

A capability is not complete for cross-platform release until both projects implement and verify its behavioral contract on their own supported host. A documented platform gap remains `NOT AT PARITY`; it must not be presented as a completed feature. Do not remove or permanently waive a capability in one sibling without the user's explicit decision recorded in both repositories.

## Shared capability baseline

WordPlai's current user-facing baseline is documented in [`modules/word/README.md`](../modules/word/README.md) and [`for-ai/PROJECT.md`](./PROJECT.md). Map every behavior there to a Writer-native equivalent. The baseline includes:

| Capability area | Required equivalent outcome |
| --- | --- |
| Local passage rules | Users mark passages and locally own color/name, prompt, and writing-function rules. Neutral marks alone never authorize wording changes. |
| Protection and review rules | Preserve-text, reject, discard, prompt editing, and overlapping-rule precedence remain explicit and restrictive where applicable. |
| AI highlighting and recoloring | AI may add neutral marks, extend an unchanged local rule by exact identity, and recolor selected marks without changing rule definitions; ambiguous or unsafe batches are refused. |
| Complete document context | AI can list/read the chosen open document completely, including relevant comments, replies, and revision context, before proposing changes. |
| Document choice and isolation | Multiple documents can be selected by name; each conversation retains its own target; closed or stale targets never redirect work to another document. |
| Authorized wording changes | Edits stay inside user-authorized prose and applicable paragraph/structure limits; completed edits include a concise rationale and protected completion marking. |
| Native review and undo | Host-native tracked changes, comments, undo/redo, document protection, and user review retain their semantics; uncertain or partial outcomes are surfaced and safely resolved. |
| Save and persistence | User rules and document marks survive ordinary save/reopen; automatic save behavior and failure reporting have an equivalent user outcome. |
| AI connection and MCP behavior | Connection setup, start/stop or revoke, public endpoint sharing, multiple clients/conversations, ordered operations, stale-context refusal, and safe retry behavior have equivalent outcomes. |
| Scope and safety boundaries | No remote file, shell, settings, or approval operations are introduced. Unsupported structures and host limits are refused clearly without flattening protected content. |

This table defines product outcomes, not a shared API schema. Keep exact host-specific protocols in each project's own source, tests, and task-routed control-plane documents.

## Change and release rules

1. Before completing a new user-visible capability in either product, identify the matching behavior and owner in the sibling project and update this contract in both repositories if the shared baseline changes.
2. Implement and validate the behavior using each host's native APIs, storage, UI, and transaction model. A Word check does not qualify Writer, or vice versa.
3. Record a per-capability status in each project's verification evidence. Mark an unimplemented, unverified, blocked, or host-limited behavior honestly; do not infer parity from similar names or screenshots.
4. Treat changes to permissions, protection, document targeting, revision/comment handling, persistence, or retry semantics as compatibility changes requiring review on both platforms.
5. Keep public descriptions and release notes aligned. Neither project may claim parity while a shared capability is missing or unverified on its supported host.

## Project links

- [WordPlai repository](https://github.com/GeorgeFejer91/wordplai)
- [LibrAI repository](https://github.com/GeorgeFejer91/LibrAI)
