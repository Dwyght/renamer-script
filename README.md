# Rename files and folders from Excel

Install the dependency, preview, then apply:

```powershell
python -m pip install -r requirements.txt
python rename_from_excel.py --config config.json --dry-run
python rename_from_excel.py --config config.json --apply
```

The console prints `Working... please wait.` immediately and repeats an activity
message every five seconds until the script finishes.

For another dataset, edit `config.json`, set your root directory and Excel
path, and select the source and destination column names. Paths are relative to
the configuration file. `sheet_name: null` selects the active worksheet. The
existing `config.json` still points to your current dataset and workbook.

A mapping `A -> B` renames `A/A.pdf` and `A/A.png` to `B/B.pdf` and `B/B.png`.
File mappings match the entire stem, preserving the extension and its casing.
Only listed extensions are renamed; other files keep their names. Folder names
match the entire mapping key. Enable both `rename_files` and `rename_folders`
to rename a folder and its matching files. Nested folders are processed first.
Symbolic links and junctions are skipped. The selected root itself is not renamed.

Boolean settings must be JSON `true` or `false`. Dry run is the default;
`--dry-run` forces preview even if the configuration enables apply. Column names
are matched without regard to case and surrounding whitespace. When omitted,
column names default to `New` and `Equivalent from old` for existing workbooks.
Mapping basenames must meet Windows filename rules on every platform.

The script checks every group for collisions and stops on traversal errors
before applying changes. Any collision prevents the entire run from applying.
Directory listings are reused for collision checks. File identities are read
only for matching rename candidates and reused when recording history.
Swaps and cycles use temporary names. If a rename fails, the script attempts to
restore the entire run, including files inside folders renamed earlier.

Applied runs keep history in `.rename-history` beside the Python script, outside
the dataset. Each dataset has a separate JSON history file named using a hash of
its absolute root path. Set `history_directory` in the configuration to use a
different location; relative paths are resolved against the configuration file,
and the history directory must be outside the dataset. The history records
filesystem identities for each mapping so a second run skips previously renamed
objects, even when a destination also appears as a source. New objects can still
be processed. Preserve this history; deleting it, copying files to a
different filesystem, or changing the mapping removes that protection for those
objects. It cannot identify renames made before history tracking was installed.
The filesystem must provide stable file identities.

A lock file in the history directory prevents concurrent apply runs. Existing
`.rename-state.json` history from older versions is moved out of the dataset on
the next apply run; previews can read it without moving it. A pending transaction in
the history blocks subsequent runs after an interruption or failed recovery.
It lists each group's original folder and `(source, temporary, destination)`
moves. If recovery is needed, preserve the history and remaining temporary files,
inspect which moves completed, and restore the original names before clearing
`pending`. Parent folders must be restored before restoring their children.
Remove a leftover lock only after verifying that no run is active. Automatic
rollback covers handled rename errors; process termination and persistent
filesystem failures may require this manual recovery.

Exit codes: `0` success, `1` collisions or a rename failure that was rolled back,
`2` invalid input, traversal errors, or recovery required.

