# Outlook → Obsidian — Bulk Export Tool (C# / COM)

A small command-line tool that exports Outlook mail to Obsidian Markdown notes
**in bulk** — by day / week / month / year / whole folder, or incrementally
since the last run.

Unlike the Graph API approach, this needs **no Azure app registration and no
admin rights**: it automates the **classic Outlook for Windows** you are already
signed into, via COM. That makes it the practical option on a locked-down
corporate machine.

## Why this instead of the VBA macro

- **Bulk + incremental** export instead of one-email-at-a-time.
- **No "enable all macros"** requirement (that VBA setting weakens Outlook's
  whole macro security posture). This is a separate, optionally code-signed exe.
- Untrusted email content is escaped before it reaches the note:
  - sender/recipient display names are **YAML-escaped** (no frontmatter injection);
  - file names are sanitized (illegal chars, reserved device names, path
    separators stripped — no traversal out of the vault);
  - the body can optionally be **fenced** so embedded Dataview/Templater code
    cannot execute on open.

## Requirements

- **Classic** Outlook for Windows, signed in. *The "new Outlook" does not expose
  the COM object model and will not work.*
- Macro/COM automation not blocked by corporate policy (your old VBA macro
  working is a good sign it isn't).
- .NET Framework 4.8 (built into Windows) to run the exe; Visual Studio or
  `msbuild` on a machine with Outlook installed to build it.

## Build

```powershell
# On a Windows machine that has classic Outlook installed:
dotnet build -c Release
# or open the .csproj in Visual Studio and build.
```

The Outlook interop types are embedded into `o2o-export.exe`, so nothing extra
needs to be deployed alongside it.

> Tip: if the build complains about the COM reference version, open the project
> in Visual Studio → **Add COM Reference** → tick *Microsoft Outlook NN.0 Object
> Library*, and adjust `VersionMajor`/`VersionMinor` in the `.csproj`.

## Configure

Copy `config.example.txt` to `config.txt` (next to the exe) and set at least
`VaultPath`. See the comments in that file for every option.

## Use

```powershell
o2o-export --range week            # last 7 days
o2o-export --range month           # last 30-ish days (calendar month back)
o2o-export --range year
o2o-export --range all             # the whole folder
o2o-export --range 2025-01-01..2025-12-31   # explicit dates (end inclusive)
o2o-export --since-last-export     # everything new since the previous run
o2o-export --range week --dry-run  # preview without writing files
```

### How "since last export" works

The tool stores a high-water mark (newest message exported) plus the set of
already-exported message ids in `state.txt`. `--since-last-export` exports only
mail newer than that mark, and the id set makes overlapping ranges idempotent —
re-running never produces duplicate notes. Delete `state.txt` to start over.

Schedule `o2o-export --since-last-export` in Windows Task Scheduler for a
hands-off daily sync (Outlook must be running when it fires).

## Limitations

- Classic Outlook for Windows only (COM); not Mac / web / new Outlook.
- Incremental tracking is timestamp + id based (Outlook COM has no server-side
  delta token). For change tracking that also captures moves/deletes you'd need
  the Graph API path, which requires an Azure app registration.
