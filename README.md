# MAPICommander 0.6 Beta

MAPICommander is a Windows command-line tool for counting, listing, and removing email from a single Exchange mailbox folder. It uses the existing classic Outlook profile through Extended MAPI, not EWS, and does not require an Entra app registration.

> **Beta software.** MAPICommander is provided **as is**, without warranty of any kind, express or implied. There is no guarantee that it will work correctly or that it is suitable for your environment. Test it on a disposable folder and review a read-only preview before deleting anything. Hard-deleted messages may not be recoverable.

## Download and requirements

Download the self-contained Windows executable from [GitHub Releases](https://github.com/reikor0x/MAPICommander/releases). No separate .NET runtime or DLLs are needed. Match the executable to **Outlook's** architecture: `win-x64` for 64-bit Outlook and `win-x86` for 32-bit Outlook.

- **Classic Outlook for Windows** with a working Exchange profile. The new Outlook is not supported. Run the tool in the Windows session that owns the profile.
- For a delegated or shared mailbox, grant the Outlook user the appropriate Exchange permissions (typically Full Access) and ensure the mailbox appears as a separate store in the selected profile. Use Outlook automapping or add the mailbox to the profile. Opening an unmounted mailbox by SMTP address is not supported.

## Quick start

The examples use PowerShell. Adjust the filename to match the executable you downloaded and use the store and folder names shown by your Outlook profile.

```powershell
$app = '.\MAPICommander-win-x64.exe'
& $app stores
& $app folders --profile 'Outlook' --store 'Shared Mailbox'
& $app folders --profile 'Outlook' --store 'Shared Mailbox' --folder 'Inbox'
```

Omit `--profile` to open the profile chooser. `--store` uses the displayed store name, not the mailbox email address. Ambiguous names are rejected. Folder paths are relative to the mailbox's mail root, with `\` between segments. Outlook may display localized folder names.

```powershell
$target = @('--profile', 'Outlook', '--store', 'Shared Mailbox', '--folder', 'Inbox\Archive')

# Count messages received before January 1, 2025 (UTC). No mailbox changes.
& $app scan @target --before '2025-01-01' --online

# Preview the messages a hard delete would target. No mailbox changes.
& $app scan @target --from '2018-01-01' --to '2028-12-31' --online --hard-delete --whatif > .\preview.tsv
```

## Filters and output

| Option | Behavior |
| --- | --- |
| `--from DATE` | Optional inclusive start. |
| `--before DATE` | Exclusive end. Choose this or `--to`, not both. |
| `--to DATE` | Inclusive final calendar day in UTC. Accepts a date only. |
| `--online` | Requests direct access to Exchange instead of the Outlook cache. If online opening fails, the tool does not fall back to cached mode. Without this option, the cache may not contain the full history. |
| `--list` | Lists matches as TSV. Without it, a normal scan prints only the count. |
| `--whatif` | Produces the same minimal TSV report without opening the folder for writing or applying a deletion. Works with `--delete` or `--hard-delete`. |

Dates without a time use UTC. Use `yyyy-MM-dd` (recommended) or `d/M/yyyy`. For `--from` and `--before`, you may also specify a timestamp with `Z` or an explicit offset, for example `2025-01-01T00:00:00+01:00`.

The filter uses the message's **received** timestamp, not its sent or creation timestamp. It matches standard mail (`IPM.Note` and subclasses) with a received timestamp in the selected folder only. It does not recurse or scan search folders. The TSV contains only `received_utc` and `entry_id_hex`. Entry IDs are opaque MAPI identifiers, not Internet Message-IDs or a reusable delete list. Diagnostics and timings go to stderr.

## Removing messages

After reviewing the preview, rerun the command without `--whatif` and without the report redirection. The tool scans again; the preview is not a snapshot or a lock on the mailbox.

- **`--delete`** moves matches to Deleted Items in the same mailbox. Type `SPOSTA` at the confirmation prompt.
- **`--hard-delete`** bypasses Deleted Items. Type `ELIMINA DEFINITIVAMENTE` at the confirmation prompt. Exchange retention and hold policies still apply; hard delete does not guarantee immediate server-side erasure.
- **`--yes`** skips interactive confirmation, including for hard delete. Use only when intentionally running unattended. It cannot be combined with `--whatif`.

Deletion is not atomic. An error or Ctrl+C may leave some messages already moved or deleted. Rerun a read-only scan before retrying. To validate the tool with your Outlook build and Exchange configuration, start with a test folder containing messages before, on, and after your date boundaries, and a message in a subfolder that should not match.

## Build from source

Use the .NET 10 SDK or later on the development machine:

```powershell
dotnet publish -c Release -r win-x64
dotnet publish -c Release -r win-x86
```

Each build produces a single `MAPICommander.exe` under `bin\Release\net10.0\<runtime>\publish`. Distribute the executable from the **publish** directory, not the regular build output. The embedded runtime may extract native components to a temporary location at startup.

The implementation applies the date filter with `IMAPITable::Restrict`, reads up to 512 rows at a time without opening individual messages, and moves or deletes in batches of up to 256. Measure performance against your mailbox; a speed advantage over EWS is not guaranteed. Extended MAPI does not have an official Microsoft .NET wrapper. The native calls are based on the patterns in [MFCMAPI](https://github.com/microsoft/mfcmapi) and the headers in [Microsoft's MAPI Stub Library](https://github.com/microsoft/MAPIStubLibrary).