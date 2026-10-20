# Wordplai CLI

`wordplai.exe` controls the installed Windows add-in through its existing local
COM authority. It requires desktop Word with a document open and Wordplai
connected. It never launches Word, sends mouse/keyboard input, or grants new
remote permissions. Normal Word busy, protection and uncertainty guards apply.
The native background-document scope restores the active document and selection.

From PowerShell after installation:

```powershell
$cli = "$env:LOCALAPPDATA/Programs/wrAIter Word/wordplai.exe"
& $cli help
& $cli status
& $cli documents
& $cli state --document DOCUMENT_ID
& $cli context --document DOCUMENT_ID
```

Use IDs returned by `documents`; titles and the active window never substitute
for a missing, closed or ambiguous target. With multiple Word processes, pass
`--word-pid PID`. Output is UTF-8 JSON; exit codes are 0 success, 1 native/MCP
refusal, 2 usage/input error, and 3 runtime/transport failure. Capability URLs are
redacted unless `--include-secrets` is supplied. Keep raw output private because
state/context includes document data.

## Local rules and markers

Commands `rule-save`, `mark`, `unmark`, `accept`, `reject`, and `revoke` act with
the local user's authority. They require `--document` and a current `snapshot`
from `context.result.snapshot` or `state.result.document.revision`. They share
the native policy, marking, review and autosave implementations used by Word.
There is no UI draft to save or cancel: `rule-save` explicitly commits a rule.

Write UTF-8 JSON to a file and pass `--input FILE`, or use `--input -` for stdin.
Do not place document text or credentials directly on the command line.

Example `rule.json`:

```json
{"snapshot":"CURRENT_SNAPSHOT","id":"concise","label":"Concise","color":"#AACCEE","action":"prompt","prompt":"Shorten while preserving meaning."}
```

```powershell
& $cli rule-save --document DOCUMENT_ID --input rule.json
# Read fresh context after changing a rule, then prepare mark.json.
& $cli context --document DOCUMENT_ID
```

Example `mark.json` (offsets are exact UTF-16 offsets in the main body):

```json
{"snapshot":"CURRENT_SNAPSHOT","ruleId":"concise","from":0,"to":21,"quote":"Exact current wording"}
```

```powershell
& $cli mark --document DOCUMENT_ID --input mark.json
```

`unmark`, `accept` and `reject` use `{"snapshot":"CURRENT_SNAPSHOT","id":"ID"}`.
`revoke` uses `{"snapshot":"CURRENT_SNAPSHOT"}`. Wrong quotes, unsupported native
structures, read-only documents and stale snapshots refuse without retargeting.
Local rule/mark commands do not perform AI wording edits.

## Exercise the actual MCP tools

```powershell
& $cli connect-local
& $cli mcp --method initialize --session-file session-a.json
& $cli mcp --method tools/list --session-file session-a.json
& $cli mcp --method tools/call --session-file session-a.json --input request.json
& $cli stop
```

`request.json` contains MCP parameters, e.g.
`{"name":"read_manuscript","arguments":{"manuscript_id":"DOCUMENT_ID","offset":0}}`.
The output retains the complete MCP result/error envelope. Continue every page
until the actual tool returns complete context; use its revision, context witness
and original request ID for `suggest_text` / `highlight_text`. The four existing
MCP tools enforce the same protection, selected-document, replay and uncertainty
rules. No direct-write bypass is added.

Use a distinct session file for each simulated client. The CLI never overwrites
an initialization file or silently rebinds it after connection restart. Session
files are private local credentials; keep them out of Git and shared reports.
There are no automatic write retries. If transport fails or an outcome is
uncertain, inspect the original request before attempting another write.
`connect-public` explicitly starts the configured public tunnel; `stop` revokes
access and performs normal cleanup. Access is never started implicitly.

## Automated checks from a source checkout

```powershell
./modules/word/build.ps1 -OutputDirectory artifacts/cli-build
./modules/word/tests/test-cli.ps1 -CliPath artifacts/cli-build/wordplai.exe -OutputDirectory artifacts/cli-checks
./modules/word/tests/test-native-review.ps1 -CliOnly -OutputDirectory artifacts/native-cli-checks
```

The first checks CLI parsing/JSON/exit behavior without Word input. The second
uses only a separately spawned, PID-bound Word fixture and real native policy;
it checks background target isolation and selection/foreground preservation.
These checks do not qualify visual rendering, physical gestures or real WebGPT.
The CLI is Windows-only; Mac implementation remains in the Mac agent handoff.
