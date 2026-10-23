# Optional Secret Tunnel Lite package

The download page offers Wordplai, Secret Tunnel Lite, and both in one installer.
Wordplai is independently usable. Lite adds a compact Connection page with zrok
registration, GitHub authorization and file-access choices, plus a Folders tab
for a main workspace and additional approved folders.

Lite reuses the canonical [SecretTunnel-v2](https://github.com/GeorgeFejer91/SecretTunnel-v2)
native backend with its `lite` Cargo feature and `tauri.lite.conf.json` build.
Its profile, local port, application identity and zrok enrollment are distinct
from the full application. No account state or folder grants are packaged.

This preview uses separate MCP URLs for Wordplai and Lite. Installing both does
not merge their tools into one endpoint.

The combined NSIS installer embeds the exact, checksum-verified standalone
installers. It preflights open Word and stops on a component failure. A successful
first component remains installed if the second fails; no rollback or user-data
deletion is attempted. The two apps retain their ordinary independent uninstallers.

Build after qualifying and recording the exact Lite source/installer:

```powershell
./modules/companion/build.ps1 -WordInstaller PATH_TO_PUBLISHED_WORD_PREVIEW8 `
  -LiteInstaller PATH_TO_QUALIFIED_LITE_INSTALLER -LiteSha256 REVIEWED_SHA256 `
  -LiteSourceRevision EXACT_40_CHARACTER_SOURCE_REVISION
```

The Word-only component is the unchanged published preview.8. Do not relabel
later Word source results as evidence for that component or vice versa. Keep
private installer qualification under ignored `artifacts/` and publish exact
hashes, source/component provenance and observed platform/clean-host limits.

Run `./modules/companion/test.ps1` to execute the actual NSIS success and failed
component fixtures, including the production preflight when Word is open.
