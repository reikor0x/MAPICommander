# Changelog

This history was reconstructed from the repository's release tags and commits. Changes not yet included in a tagged release are listed under **Unreleased**.

## Unreleased

No unreleased changes yet.

## [0.9.0 Beta](https://github.com/reikor0x/MAPICommander/releases/tag/v0.9.0-beta)

- Added experimental export of previewed messages to a new Unicode PST. Full content (bodies, recipients, and attachments) is fetched only on export, not during preview or removal.

## [0.8.1 Beta](https://github.com/reikor0x/MAPICommander/releases/tag/v0.8.1-beta)

- Added a custom mail icon to the executable, window, Windows taskbar, and version label.

## [0.8.0 Beta](https://github.com/reikor0x/MAPICommander/releases/tag/v0.8.0-beta)

### Added

- The GUI can search the selected folder, include its nested subfolders, or search the entire mailbox. It warns before a mailbox-wide search.
- **Largest first** displays message sizes and sorts results by size in any search scope. The date filter can be turned off.
- Preview results show the message subject and source folder. The TSV export does not include subjects.
- For both **Move to Deleted Items** and **Hard delete**, the confirmation offers all previewed messages or only the highlighted rows. Removal never extends beyond the chosen preview rows.
- The GUI shows the build version and offers English (default), Italian, French, German, and Spanish under **Options > Language**.

### Changed

- The active folder or mailbox remains highlighted during processing, with its path displayed above the results.
- The larger, full-width progress bar shows a percentage when a total is available. Successful searches end with **Search ended** and a match count.
- Renamed the GUI connection option to **Online mode no cache** without changing its behavior. CLI messages and interactive confirmations are now in English.

## [0.7.4 Beta](https://github.com/reikor0x/MAPICommander/releases/tag/v0.7.4-beta)

- Fixed a GUI removal path that could fail with `0x80040107` after a successful preview. Move and hard delete now reuse the preview's MAPI store and reopen the exact source folder by its EntryID.
- Removal errors now distinguish between failures before any batch was submitted and failures after a batch may have changed the mailbox.

## [0.7.3 Beta](https://github.com/reikor0x/MAPICommander/releases/tag/v0.7.3-beta)

- Kept browsed stores open and addressed `0x80040107` when opening nested folders or previewing messages. If the store rejects a folder EntryID, the app retries the same ID through its parent folder and reports both HRESULTs if both calls fail.

## [0.7.2 Beta](https://github.com/reikor0x/MAPICommander/releases/tag/v0.7.2-beta)

- Corrected the native `IMsgStore::OpenEntry` and `CompareEntryIDs` method bindings. This resolved `0x80040106` when expanding a mailbox.

## [0.7.1 Beta](https://github.com/reikor0x/MAPICommander/releases/tag/v0.7.1-beta)

- Replaced manual Outlook profile entry with a drop-down populated from the profiles available to the current Windows user.
- Fixed `0x80040103` from `GetMsgStoresTable` by no longer requesting the Unicode flag on the table call itself. Store names still prefer Unicode, with an ANSI fallback.

## [0.7.0 Beta](https://github.com/reikor0x/MAPICommander/releases/tag/v0.7.0-beta)

- Added a WinForms desktop interface to the same executable as the CLI: mailbox and folder browsing, UTC date filters, read-only preview, TSV export, and confirmed removal.
- Added GUI progress and cancellation between MAPI batches.

## [0.6.1 Beta](https://github.com/reikor0x/MAPICommander/releases/tag/v0.6.1-beta)

- Added CLI progress reporting for scans and removal. Progress goes to standard error so it does not interfere with the TSV report on standard output.

## [0.6.0 Beta](https://github.com/reikor0x/MAPICommander/releases/tag/v0.6.0-beta)

- Initial tagged release: Windows CLI using classic Outlook profiles through Extended MAPI, including delegated Exchange mailboxes mounted in the profile.
- Added folder listing, received-date filtering, count and minimal TSV output, and a read-only `--whatif` preview.
- Supported moving matches to Deleted Items or hard-deleting them, with explicit confirmation unless `--yes` was supplied.
- Shipped single-file, self-contained executables for x64 and x86 Outlook.