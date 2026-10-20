# wordplai for Microsoft Word

<img src="../../assets/wraiter/icon.svg" width="64" height="64" alt="wordplai AI waiter and typewriter logo">

The approved icon is shared across the ribbon, add-on windows, setup/uninstall
and Windows installed-app entry. [Download wordplai](https://georgefejer91.github.io/wordplai/download/).
Dialogs use its cream and deep-green palette, with a small logo beside the
bottom actions. Controls keep their native behavior; high contrast uses system
colors.

A native Windows Word add-in with a single installer. Select prose, then click
a color square on Word's **Home** or **wordplai** ribbon to mark it. Double-click
the square or an already marked passage to open that color's rule dialog.
Rule windows keep Word available for another passage double-click. Different
colors open their own matching windows; returning to a color retains its unsaved
draft. A changed passage can still open its rule through its surviving native
marker. Current native bookmark ranges, visible colors and saved local rules
automatically reconcile ordinary local wording changes and recover valid stale
markers. Missing identities, removed colors, changed paragraph boundaries and
ambiguous native review/Undo states remain blocked with a specific diagnostic.
Shading alone never grants permission. Overlapping markers require an
explicit choice. The dialog places its logo at the top left and shows prompt
presets in three columns. Wording edits apply directly; there is no mode toggle.
Only Save color rule applies the chosen function or instructions.
Color selections may span paragraphs, include their paragraph marks and contain
manual soft line breaks. AI wording edits still stay within one authorized paragraph.
Overlapping colors offer an explicit choice. Ribbon keytips provide keyboard
access. Each of up to 24 colors has its own dialog; there is no central highlight
or suggestion window. **New color rule** adds a color; **Connections** manages
AI access. No separate application or tray utility is needed. The installed
add-in owns its internal tunnel process and private connection profile.

## Command-line automation

Source builds and preview.5 packaging include `wordplai.exe` and
[CLI instructions](cli/README.md). The CLI provides JSON commands for document
discovery, explicitly targeted state/context, local rules and UTF-16 range
marking, native review, connection lifecycle and the existing MCP tools. It uses
the add-in's local authority without mouse/keyboard input or a separate service.
Rule changes commit explicitly; CLI behavior checks do not replace visual or
physical interaction evidence. Preview.3 installers do not contain this CLI.
Preview.5 enrolls successful local changes to inactive documents for native
autosave before any MCP read, without sharing remote editing authority.

Requires installed desktop Word on Windows 10/11 x64 (Intel/AMD) and .NET Framework 4.8 (included
in supported current Windows installations). The installer registers only this
add-in for the current Windows user. Close Word before installing/updating.

Each color dialog offers a name, native color picker, **Preserve text**,
wording presets and a custom prompt. Shorten and More detail are mutually
exclusive. Preserve text disables conflicting controls and prevents AI wording
changes even when an overlapping color requests edits. Fresh neutral marks grant
no editing permission until you assign a function locally. Existing protect,
reject and discard rules retain their restrictive meanings; ordinary Word
shading is not itself an AI permission.

Ask the AI to highlight text, such as all citations in green, including text
that has not been marked before. `highlight_text` accepts optional
`marking.color` as `#RRGGBB`: an unused RGB creates a neutral color; a unique
existing RGB applies its existing color and rule unchanged. An exact
`marking.colorId` can extend any existing color, including Preserve text,
Reject, Discard and prompt colors, including legacy saved rules. If several identities
share an RGB, the exact color identity is required. The AI cannot define or
change the rule or prompt; use the local double-click dialog for that.
Coloring alone edits no wording. Protection and visible-color overlap checks
refuse a batch that would hide its mark or change protected native formatting.
You can also ask to recolor existing marked passages. The AI supplies each
current marker's exact range and ID in `marking.recolorMarkIds`; the whole batch
replaces those selected bindings with the chosen color. Existing rule definitions
remain unchanged, and unselected protection still applies. A fresh target RGB
remains neutral. Partial ranges, stale identities, unresolved tracked reviews
and recolors that would discard a saved passage prompt are refused with a
specific reason. Word Undo/Redo restores the corresponding visible colors and
bindings; later local rule decisions retain priority.
For example, changing green `#00FF00` marks to unused red `#FF0000` keeps the
text and exact ranges and creates neutral red with an empty prompt. Red needs
no editing prompt. If red already has a local rule, that rule is reused
unchanged; a color name or RGB never assigns a function by itself.
Repeated or overlapping marks can be recolored together by supplying every
affected marker's exact range and ID, sorted by range start.
MCP coloring verifies the actual character shading before confirming success;
an unverified paint is rolled back. Mixed inherited and explicit shading is
checked through exact subranges, with a per-operation budget of 256 additional
probes and one second for inconclusive checks before a precise refusal.
These rules apply within this Word add-in.

**Wordplai uses direct editing for authorized prompt wording.** Each finished
paragraph changes immediately with a short rationale comment and green
Preserve text protection. Legacy prompt rules work in the same fixed edit mode;
reading or extending their colors leaves their saved definitions unchanged.
Existing tracked changes and Discard proposals still use Word's native Review
controls. A tracked replacement may appear as a deletion/insertion pair. AI should
replace an entire paragraph when the paragraph is fully authorized; a partial
highlight authorizes only that fragment. Replacements receive green
Preserve text markers after Word confirms
the complete replacement was accepted; rejected or partially reviewed changes
do not. Existing protected native units retain their identities and formatting.

Add Word comments to a paragraph to give feedback. The next complete
context read includes comments, replies and revision information for the AI to
use with the color's instructions. Changing feedback invalidates an earlier
context read. Resolve pending tracked changes locally before a new direct
wording request in that passage. The add-in's rationale is anchored to the
revised passage. For comments affected by a direct replacement, a shared ending
must remain inside each affected comment and outside protected units. Unsupported anchors are
refused before editing; native comment threads are never recreated.
Accepting a complete deletion in Word may remove comments attached to that text.

Connections contains zrok setup, Start access, Stop / revoke, a copyable public
MCP URL. Connecting authorizes complete reading of open Word documents,
including unsaved text, without a separate sharing tab or confirmation. Initial public setup requires a
zrok account enrollment token and internet. Marking, prompts and local review
work offline. ChatGPT, Claude and public tunnels require internet; a local AI model is
not included. Stop / revoke ends access;
reconnecting makes open documents readable again. Highlight and color rules
continue to control editing permissions.

With one open document, the connection reads it immediately. With several open
documents, ChatGPT presents a named document picker and waits for one choice.
It keeps that document until you explicitly switch, even when another Word
window is active. A closed target never redirects edits to another document.

The same MCP URL accepts simultaneous ChatGPT, Claude and multiple GPT
connections. Each transport initializes a standard `Mcp-Session-Id` and sends
it on later requests. A valid `openai/session` hint retains the logical
conversation's document choice, paginated reads, complete-context witnesses and
request identities across fresh transports. Different hints remain separate;
missing hints use only the physical session's default scope. Without an inline document picker,
choose a document by its returned name in chat.
The server assigns monotonically increasing FIFO tickets before native Word
operations; clients need no synchronized clock or explicit lock tools. Only one
native turn runs at a time. A queued proposal using an older revision is refused
with current context, so it cannot overwrite the preceding client's change.
Use the same request ID only for an exact network retry; a recomputed proposal
after a definite refusal gets a new ID. An uncertain running operation retains
its turn until it finishes and must be inspected before retrying. There are at
most 512 physical sessions, 64 labelled conversations and eight queued/running
calls. Both registries have independent 30-minute idle limits and include
pending native cleanup in their capacity bounds. Physical idle expiry retires
its queued work and default context while a live logical conversation survives.
Logical expiry retires its context and queued calls. `DELETE` also retires every
logical conversation that the named physical session used, including calls on
other transports; Stop retires all state. Sessions and hints organize callers and do not add permissions or
authenticate a provider account. Actual Claude/ChatGPT public qualification and
installed-package evidence remain separate from local protocol tests.

The connection instructs the AI to read the entire chosen document immediately,
understand its context and immediately carry out the requested writing task.
Each completed paragraph is sent to Word promptly while writing continues.
If no task was given, it should ask one short contextual writing question. Routine
tool choices and permission checks stay out of the conversation. These are
client instructions; actual AI behavior requires a connected-client check.

Say **"use wordplai MCP"** to start editing. The connection identifies itself as
your writing assistant in Word and treats this short request as the instruction
to follow the prompts already attached to the marked passages. It reads the
document and sends each finished paragraph to Word without asking you to repeat
the task. The Word integration was previously named wrAIter for Word; existing
installations retain their documents, local rules and connection setup.

The remote tools are `list_manuscripts`, `read_manuscript`, `suggest_text` and
the color-only `highlight_text`. Complete current document context must be read sequentially
before proposals or highlighting. Wording proposals use UTF-16 offsets inside
authorized highlighted main-body prose; highlighting may cover unmarked prose.
Headers, footnotes and other stories are read-only context. Stale snapshots,
boundary overruns and revoked grants reject remote changes; unmarked gaps and
protection also restrict wording proposals.
Small-document edits return the complete updated context for the next paragraph,
avoiding an extra read and verification round trip. Longer documents still use
complete sequential reading. Each conversation retains its own chosen document;
clients without conversation metadata use only their own initialized session's default choice.
Selected background documents also refresh and save automatically.
No remote file, shell, settings, save or approval operation exists.

The native add-in refreshes writing permissions and automatically saves the
active named document after a short idle period. New documents still use Word's
first Save As. Busy native Undo operations, read-only documents and document
protection defer automatic work; save failures are shown and retried. Reads and
edit tools do not themselves invoke Save. A suggestion and its rationale share
one Word Undo record; green marking after native acceptance has its own
Undo step. Passage rules, bookmarks, native revisions and comments persist with
Word's ordinary save, including the automatic save. Request receipts
are stored encrypted for the current Windows user, independently from the
document. Uncertain outcomes cannot be replayed as a fresh write.

Selective replacements edit prose within a paragraph, outside tables,
equations, embedded objects and locked controls. Unsupported revision overlap
and mixed character formatting are rejected rather than flattened. Protected
fields require verified native preservation. These structures remain available as reading
context. Word on Mac and browser-based Word are not supported by this package.
Tracked enclosing rewrites preserve whole protected native units through
explicit `replacementParts`: use each `protectedId` exactly once. Separate units may
move; continuous units stay intact with their original fields, links and formatting.
Native Word Accept/Reject and Undo/Redo retain their identities. Resolve partial
native review states locally before requesting another change to that paragraph.
Bookmark protection requires inspecting the document's Word XML; exports above
64 MiB are refused, including documents whose embedded images exceed that limit.

Build on Windows x64 with PowerShell 7, Node.js 22.13 or newer, .NET Framework 4.8
and NSIS 3. No separate application checkout is required:

```powershell
npm ci
npm run widget:build
./modules/word/build.ps1 -Installer -NsisExecutable C:/tools/NSIS/makensis.exe
```

If `makensis.exe` is on PATH, omit `-NsisExecutable`. The installer build downloads
the pinned official zrok 2.0.4 Windows archive and verifies both its archive and
executable SHA-256 before packaging. For an offline build, supply the verified
executable with `-ZrokExecutable C:/path/to/zrok2.exe`. No account profile or
enrollment state is copied into the installer. A DLL-only build omits `-Installer`
and needs no tunnel download or NSIS.

The build embeds the repository's local selective-editor contract, keeps legacy
COM/assembly/profile identities for upgrades, and records source/payload SHA-256
identities. Outputs are `artifacts/word/build/` and `artifacts/word/wordplai-setup.exe`.
Generated diagnostics and native qualification documents belong under ignored
`artifacts/word/`, never in source control. Node is build/test tooling, not an
installed add-in dependency. Existing editing and client limitations are recorded
in the repository's [handoff](https://github.com/GeorgeFejer91/wordplai/blob/main/for-ai/HANDOFF.md).

Run `./modules/word/test.ps1` for source-bound policy, COM ABI, actual loopback
transport and native control checks. Its report explicitly leaves installed
Word and public connectivity unqualified. Add `-NativeWord` for the isolated
real-Word tracked-revision and comment-feedback tests on disposable documents.
Add `-InstalledWord` only after Word
is activated and opened normally; that gate uses its installed add-in and a
synthetic DOCX, leaving the demonstration document open. Public zrok/AI-client
verification remains a distinct gate.
