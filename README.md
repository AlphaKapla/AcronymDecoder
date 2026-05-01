# Acronym Lookup — Windows (Rust)

A small background tool. Hover the mouse over any word, press a hotkey,
and a popup shows the definition from your personal `acronyms.csv`.

| Hotkey | Action |
|---|---|
| **Ctrl+Shift+A** | Look up the word at the cursor |
| **Ctrl+Shift+E** | Open the CSV file in your default editor |
| **Ctrl+Shift+Q** | Quit |

## The CSV file

Two columns: acronym, definition. The program finds the file in this order:

1. A path passed as a command-line argument.
2. A file called `acronyms.csv` next to the `.exe` (standard location for a distributed binary).
3. The `acronyms.csv` at the repository root, baked in at compile time — used automatically when you run the program with `cargo run`, so that file is the single source of truth during development.

To override with a custom path:

```powershell
acronym-lookup.exe "C:\Users\me\Documents\my-acronyms.csv"
```

Format rules:

* First column is the acronym, second column is the definition.
* Quoted fields are supported, so definitions may contain commas:
  `RAG,"Retrieval-Augmented Generation, grounded in documents"`
* Lookup is case-insensitive — `api`, `Api` and `API` all match.
* A header row whose first cell is `acronym`, `term`, `key`,
  `abbreviation`, or `abbr` (any case) is auto-skipped.
* **Multiple definitions for the same acronym are allowed.** Repeat the
  acronym on as many rows as you like; all of them will be shown when
  you look it up. The included starter file does this for `MVP` and
  `PR` — useful for context-dependent abbreviations.
* The file is reloaded on every press, so you can edit and save while
  the tool runs and the next press uses the updated definitions.

## Smart lookup

If the exact term isn't in the file, two fallbacks kick in:

1. **Substring match.** Catches plurals and inflected forms — `APIs`
   finds `API`, `OKRs` finds `OKR`, etc.
2. **Levenshtein typo match.** Catches near-misses — `RIO` suggests
   `ROI`, `MVPP` suggests `MVP`. The threshold tightens for short
   queries so similar 3-letter acronyms don't bleed together.

When fallbacks fire the popup says *"'X' is not in your file. Did you
mean:"* and lists up to five suggestions with their definitions. If
nothing matches at all, you get a hint to press Ctrl+Shift+E to add it.

## Edit hotkey

**Ctrl+Shift+E** opens the CSV in whatever app Windows has registered
as the default for `.csv`. That's usually Excel; if you've used "Open
With → Always use this app" to set Notepad or VS Code as your
preferred editor for `.csv`, that's what opens.

If the file doesn't exist yet, the tool creates it (with a header row)
before opening — so the very first edit press gives you somewhere to
start typing.

## Build and run

You need a Rust toolchain (`rustup` from <https://rustup.rs>).

```powershell
cargo run --release
```

That gives you a binary at `target/release/acronym-lookup.exe`. Copy
the binary and your `acronyms.csv` to wherever you want; leave the
console window open while you use it.

## How it works

* **Global hotkeys** — `RegisterHotKey` with a NULL `HWND` posts
  `WM_HOTKEY` directly to the calling thread's message queue. The main
  thread runs a stock `GetMessageW` / `DispatchMessageW` loop. No
  window is created at the top level.
* **Reading the word at the cursor** — Microsoft UI Automation:
  `IUIAutomationTextPattern::RangeFromPoint` expanded to the enclosing
  word. Two fallbacks (`CurrentName`, `ValuePattern`) cover buttons,
  menu items and edit boxes.
* **Lookup** — every press re-reads the CSV with the `csv` crate into
  a `HashMap<String, Vec<String>>`. Three-stage search: exact →
  substring → Levenshtein.
* **Popup** — built with `native-windows-gui`. Topmost, wraps long
  definitions, scrollable, closes on Esc or the Close button.
  Spamming the lookup hotkey while a popup is already open is
  ignored (one popup at a time).
* **Edit hotkey** — `ShellExecuteW` with the `open` verb on the CSV
  path, falling through to the user's default-handler association.

## Caveats

* Windows 10 / 11 only.
* Apps with custom-rendered text (some games, certain canvas-based web
  apps) may not expose their text via UIA. The popup will say "Nothing
  detected".
* The first build pulls down `native-windows-gui` and a few
  Win32 helpers; expect a one-time ~2 minute compile. Incremental
  builds afterwards are quick.
