//! Acronym Lookup — Windows Background Tool (Rust)
//!
//! Hotkeys:
//!   Ctrl+Shift+A   look up the selected word
//!   Ctrl+Shift+E   open the CSV file in the default editor
//!   Ctrl+Shift+Q   quit
//!
//! The CSV file:
//!   * Two columns: acronym, definition
//!   * Lives next to the .exe by default (file name: `acronyms.csv`)
//!   * Override with a CLI argument:
//!         acronym-lookup.exe "C:\path\to\my-list.csv"
//!   * Reloaded on every press, so edits show up immediately
//!   * Same acronym can appear on multiple rows — all definitions are
//!     shown together
//!
//! Lookup strategy (first match wins):
//!   1. Exact match (case-insensitive) on the uppercased term
//!   2. Substring match — keys that contain the query, or queries that
//!      contain a key (≥3 chars). Catches plural / inflected forms
//!   3. Levenshtein-distance match — distance ≤1 for short queries,
//!      ≤2 for queries of length ≥6. Catches typos
//!
//! Word detection strategy (first success wins):
//!   1. UI Automation TextPattern selection — reads the currently selected
//!      text from the focused element's accessibility API
//!   2. Clipboard simulation — sends Ctrl+C to the foreground window, reads
//!      the clipboard, then restores the previous clipboard content. Works
//!      with applications that don't expose a TextPattern (e.g. Adobe
//!      Acrobat). If nothing is selected in the target app, the clipboard is
//!      left unchanged and no lookup is performed.
//!
//! Build & run:
//!     cargo run --release
//!     cargo run --release -- "C:\path\to\my-list.csv"

#![windows_subsystem = "console"]

use std::collections::HashMap;
use std::env;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, AtomicU32, Ordering};
use std::sync::OnceLock;

use windows::core::{Interface, PCWSTR};
use windows::Win32::Foundation::{BOOL, GlobalFree, HANDLE, HGLOBAL, HINSTANCE, HWND, LPARAM, LRESULT, POINT, RECT, WPARAM};
use windows::Win32::Graphics::Gdi::{GetStockObject, HBRUSH, DEFAULT_GUI_FONT};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::System::Com::{
    CoCreateInstance, CoInitializeEx, CLSCTX_INPROC_SERVER, COINIT_APARTMENTTHREADED,
};
use windows::Win32::System::Console::{
    SetConsoleCtrlHandler, CTRL_BREAK_EVENT, CTRL_C_EVENT, CTRL_CLOSE_EVENT,
};
use windows::Win32::System::DataExchange::{
    CloseClipboard, EmptyClipboard, GetClipboardData, GetClipboardSequenceNumber,
    OpenClipboard, SetClipboardData,
};
use windows::Win32::System::Memory::{GlobalAlloc, GlobalLock, GlobalUnlock, GMEM_MOVEABLE};
use windows::Win32::System::Threading::GetCurrentThreadId;
use windows::Win32::UI::Accessibility::{
    CUIAutomation, IUIAutomation, IUIAutomationTextPattern,
    UIA_TextPatternId,
};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    RegisterHotKey, SendInput, UnregisterHotKey,
    INPUT, INPUT_0, INPUT_KEYBOARD, KEYBDINPUT, KEYBD_EVENT_FLAGS, KEYEVENTF_KEYUP,
    MOD_CONTROL, MOD_SHIFT, VIRTUAL_KEY, VK_CONTROL,
};
use windows::Win32::UI::Shell::ShellExecuteW;
use windows::Win32::UI::WindowsAndMessaging::{
    AdjustWindowRectEx, CreateWindowExW, DefWindowProcW, DestroyWindow, DispatchMessageW,
    ES_AUTOVSCROLL, ES_MULTILINE, ES_READONLY, GetCursorPos, GetMessageW, IDC_ARROW,
    IsDialogMessageW, LoadCursorW, MB_ICONERROR, MB_OK, MB_TOPMOST, MessageBoxW, MSG,
    PostQuitMessage, PostThreadMessageW, RegisterClassExW, SendMessageW, SW_SHOWNORMAL,
    TranslateMessage, WINDOW_EX_STYLE, WINDOW_STYLE, WNDCLASSEXW, WM_CLOSE, WM_COMMAND,
    WM_DESTROY, WM_HOTKEY, WM_QUIT, WM_SETFONT, WS_CAPTION, WS_CHILD, WS_EX_CLIENTEDGE,
    WS_EX_TOPMOST, WS_SYSMENU, WS_VISIBLE, WS_VSCROLL, HMENU,
};

// ----- configuration --------------------------------------------------------
const HOTKEY_LOOKUP: i32 = 1;
const HOTKEY_QUIT:   i32 = 2;
const HOTKEY_EDIT:   i32 = 3;
const VK_A: u32 = 0x41;
const VK_E: u32 = 0x45;
const VK_Q: u32 = 0x51;

const CSV_FILE_NAME: &str = "acronyms.csv";

// Suggestion tuning
const SUGGESTION_LIMIT: usize = 5;
const SUBSTRING_MIN_LEN: usize = 3;

// ---------------------------------------------------------------------------
// A single popup-in-flight at a time. The flag is cleared from the
// popup thread once the window closes, so spamming the hotkey doesn't
// stack windows on top of each other.
// ---------------------------------------------------------------------------
fn popup_busy() -> &'static AtomicBool {
    static FLAG: OnceLock<AtomicBool> = OnceLock::new();
    FLAG.get_or_init(|| AtomicBool::new(false))
}

// ---------------------------------------------------------------------------
// Console-control handler — intercepts Ctrl+C, Ctrl+Break and console-close
// events so that they trigger a clean shutdown (exit code 0) instead of the
// Windows default ExitProcess(STATUS_CONTROL_C_EXIT) which cargo reports as
// "process didn't exit successfully".
//
// The handler runs on a thread created by Windows; we post WM_QUIT to the
// main thread so the regular message loop exits and does its cleanup.
// ---------------------------------------------------------------------------
fn main_thread_id() -> &'static AtomicU32 {
    static ID: OnceLock<AtomicU32> = OnceLock::new();
    ID.get_or_init(|| AtomicU32::new(0))
}

unsafe extern "system" fn console_ctrl_handler(ctrl_type: u32) -> BOOL {
    match ctrl_type {
        CTRL_C_EVENT | CTRL_BREAK_EVENT | CTRL_CLOSE_EVENT => {
            let tid = main_thread_id().load(Ordering::Relaxed);
            if tid != 0 {
                // Best-effort: ignore the return value.
                let _ = PostThreadMessageW(tid, WM_QUIT, WPARAM(0), LPARAM(0));
            }
            BOOL(1) // handled — do NOT call the default handler
        }
        _ => BOOL(0), // not handled — pass to the next handler
    }
}

// ----- entry point ----------------------------------------------------------
fn main() -> windows::core::Result<()> {
    let path = csv_path();

    println!("============================================================");
    println!("  Acronym Lookup running in the background.");
    println!("  Lookup : Ctrl+Shift+A   (select a word, then press)");
    println!("  Edit   : Ctrl+Shift+E   (open the CSV in your editor)");
    println!("  Quit   : Ctrl+Shift+Q");
    println!("  CSV    : {}", path.display());
    match load_acronyms(&path) {
        Ok(map) => {
            let total: usize = map.values().map(|v| v.len()).sum();
            println!("  Loaded {} unique acronym(s), {} definition(s).", map.len(), total);
        }
        Err(e) => println!("  WARNING: {}", e),
    }
    println!("============================================================");

    // NWG runs on whichever thread calls nwg::init / nwg::dispatch_thread_events;
    // we always create popups on a fresh thread, so init happens there.

    unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);

        // Store the main thread ID and register a console-control handler so
        // that Ctrl+C / Ctrl+Break / console-close cause a clean exit (code 0)
        // rather than STATUS_CONTROL_C_EXIT.
        main_thread_id().store(GetCurrentThreadId(), Ordering::Relaxed);
        let _ = SetConsoleCtrlHandler(Some(console_ctrl_handler), BOOL(1));

        RegisterHotKey(HWND::default(), HOTKEY_LOOKUP, MOD_CONTROL | MOD_SHIFT, VK_A)?;
        RegisterHotKey(HWND::default(), HOTKEY_EDIT,   MOD_CONTROL | MOD_SHIFT, VK_E)?;
        RegisterHotKey(HWND::default(), HOTKEY_QUIT,   MOD_CONTROL | MOD_SHIFT, VK_Q)?;

        let mut msg = MSG::default();
        while GetMessageW(&mut msg, HWND::default(), 0, 0).0 > 0 {
            if msg.message == WM_HOTKEY {
                match msg.wParam.0 as i32 {
                    HOTKEY_QUIT => break,
                    HOTKEY_LOOKUP => {
                        std::thread::spawn(do_lookup);
                    }
                    HOTKEY_EDIT => {
                        std::thread::spawn(open_csv_in_editor);
                    }
                    _ => {}
                }
            }
            let _ = TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }

        let _ = UnregisterHotKey(HWND::default(), HOTKEY_LOOKUP);
        let _ = UnregisterHotKey(HWND::default(), HOTKEY_EDIT);
        let _ = UnregisterHotKey(HWND::default(), HOTKEY_QUIT);

        // Remove the console-ctrl handler before we exit.
        let _ = SetConsoleCtrlHandler(Some(console_ctrl_handler), BOOL(0));
    }
    Ok(())
}

// ----- per-press worker -----------------------------------------------------
fn do_lookup() {
    if popup_busy().swap(true, Ordering::AcqRel) {
        // A popup is already open; ignore this press.
        return;
    }

    let raw  = unsafe { get_text_at_cursor() };
    let term = raw.as_deref().and_then(extract_candidate_acronym);

    let result = match term {
        Some(t) => {
            let r = lookup(&t);
            (t, r)
        }
        None => (
            "Nothing detected".to_string(),
            LookupResult::Error(
                "Could not detect any selected text.\n\n\
                 Select a word (double-click or click-and-drag) and try again.\n\n\
                 The tool first reads the selection via the Windows accessibility \
                 API (UI Automation). If that fails it retries by sending Ctrl+C \
                 and reading the clipboard. Some applications (certain games, \
                 custom-rendered canvases) support neither method."
                    .to_string(),
            ),
        ),
    };

    show_popup(&result.0, result.1);
    popup_busy().store(false, Ordering::Release);
}

// ----- open the CSV in the default editor ----------------------------------
fn open_csv_in_editor() {
    let path = csv_path();

    if !path.exists() {
        // Create an empty file with a header row so editing has somewhere
        // to start. This keeps the "edit" hotkey useful on first run.
        if let Err(e) = std::fs::write(&path, "acronym,definition\n") {
            error_box(&format!(
                "Could not create '{}':\n\n{}",
                path.display(),
                e
            ));
            return;
        }
    }

    // ShellExecuteW with verb "open" runs the file's default-handler
    // association. For .csv that's usually Excel, but the user's
    // preference (e.g. Notepad, VS Code) is honoured if they've changed
    // it via Open With → "Always use this app".
    let path_w: Vec<u16> = path
        .as_os_str()
        .encode_wide()
        .chain(std::iter::once(0))
        .collect();
    let verb_w: Vec<u16> = "open\0".encode_utf16().collect();

    unsafe {
        let result = ShellExecuteW(
            HWND::default(),
            PCWSTR(verb_w.as_ptr()),
            PCWSTR(path_w.as_ptr()),
            PCWSTR::null(),
            PCWSTR::null(),
            SW_SHOWNORMAL,
        );
        // Per docs, ShellExecuteW returns >32 on success.
        if (result.0 as isize) <= 32 {
            error_box(&format!(
                "Could not open '{}' in the default editor.",
                path.display()
            ));
        }
    }
}

use std::os::windows::ffi::OsStrExt;

// ----- where the CSV lives --------------------------------------------------
fn csv_path() -> PathBuf {
    if let Some(arg) = env::args().nth(1) {
        return PathBuf::from(arg);
    }
    env::current_exe()
        .ok()
        .and_then(|p| p.parent().map(|d| d.join(CSV_FILE_NAME)))
        .unwrap_or_else(|| PathBuf::from(CSV_FILE_NAME))
}

// ----- CSV loader -----------------------------------------------------------
//
// Returns a map from uppercased acronym -> Vec<definition>. Multiple rows
// with the same key are preserved as separate entries.
fn load_acronyms(path: &Path) -> Result<HashMap<String, Vec<String>>, String> {
    let mut reader = csv::ReaderBuilder::new()
        .has_headers(false)
        .flexible(true)
        .trim(csv::Trim::All)
        .from_path(path)
        .map_err(|e| format!("Could not open '{}': {}", path.display(), e))?;

    let mut map: HashMap<String, Vec<String>> = HashMap::new();
    for (idx, result) in reader.records().enumerate() {
        let record = match result {
            Ok(r) => r,
            Err(_) => continue,
        };
        if record.len() < 2 { continue; }

        let key   = record[0].to_ascii_uppercase();
        let value = record[1].to_string();
        if key.is_empty() || value.is_empty() { continue; }

        if idx == 0 && matches!(
            key.as_str(),
            "ACRONYM" | "TERM" | "KEY" | "ABBREVIATION" | "ABBR"
        ) {
            continue;
        }

        map.entry(key).or_default().push(value);
    }
    Ok(map)
}

// ----- lookup ---------------------------------------------------------------
enum LookupResult {
    /// Exact match: `(key, all definitions)`
    Exact(String, Vec<String>),
    /// No exact match, but suggestions to offer. `(query, [(key, distance, definitions)])`
    Suggestions(String, Vec<(String, usize, Vec<String>)>),
    /// No exact match and no suggestions. `(query, total entries searched)`
    NotFound(String, usize),
    /// Anything that prevented us from looking up at all (e.g. file missing).
    Error(String),
}

fn lookup(term: &str) -> LookupResult {
    let path = csv_path();
    let map = match load_acronyms(&path) {
        Ok(m) => m,
        Err(e) => {
            return LookupResult::Error(format!(
                "Could not load acronyms file.\n\n\
                 {}\n\n\
                 Expected location:\n  {}\n\n\
                 Create a CSV file with two columns (acronym, definition) at \
                 that path, or pass a different path as a command-line argument:\n  \
                 acronym-lookup.exe \"C:\\path\\to\\your-file.csv\"",
                e,
                path.display()
            ));
        }
    };

    let key = term.to_ascii_uppercase();

    // 1. Exact match.
    if let Some(defs) = map.get(&key) {
        return LookupResult::Exact(key, defs.clone());
    }

    // 2. Substring match (only meaningful for queries of decent length).
    let mut substring_hits: Vec<(String, Vec<String>)> = Vec::new();
    if key.len() >= SUBSTRING_MIN_LEN {
        for (k, defs) in &map {
            if k == &key { continue; }
            if k.contains(&key) || key.contains(k.as_str()) {
                substring_hits.push((k.clone(), defs.clone()));
            }
        }
    }

    // 3. Levenshtein for typos. Threshold tightens for short queries
    //    so that "API" doesn't match "AWS" at distance 2.
    let lev_threshold: usize = if key.len() <= 3 { 1 }
                               else if key.len() <= 5 { 1 }
                               else { 2 };

    let mut lev_hits: Vec<(String, usize, Vec<String>)> = Vec::new();
    for (k, defs) in &map {
        if k == &key { continue; }
        // Cheap length filter — distance is at least |len_a - len_b|.
        if k.len().abs_diff(key.len()) > lev_threshold { continue; }
        let d = levenshtein(&key, k);
        if d <= lev_threshold {
            lev_hits.push((k.clone(), d, defs.clone()));
        }
    }

    if substring_hits.is_empty() && lev_hits.is_empty() {
        return LookupResult::NotFound(term.to_string(), map.values().map(|v| v.len()).sum());
    }

    // Merge suggestions: substring matches get pseudo-distance 0 so
    // they sort to the top, ahead of Levenshtein matches.
    let mut combined: Vec<(String, usize, Vec<String>)> = Vec::new();
    for (k, defs) in substring_hits {
        combined.push((k, 0, defs));
    }
    for (k, d, defs) in lev_hits {
        if !combined.iter().any(|(existing, _, _)| existing == &k) {
            combined.push((k, d, defs));
        }
    }
    combined.sort_by(|a, b| a.1.cmp(&b.1).then_with(|| a.0.cmp(&b.0)));
    combined.truncate(SUGGESTION_LIMIT);

    LookupResult::Suggestions(term.to_string(), combined)
}

// ----- Levenshtein (iterative, two-row) -------------------------------------
fn levenshtein(a: &str, b: &str) -> usize {
    if a == b { return 0; }
    let a: Vec<char> = a.chars().collect();
    let b: Vec<char> = b.chars().collect();
    if a.is_empty() { return b.len(); }
    if b.is_empty() { return a.len(); }

    let mut prev: Vec<usize> = (0..=b.len()).collect();
    let mut curr: Vec<usize> = vec![0; b.len() + 1];

    for i in 1..=a.len() {
        curr[0] = i;
        for j in 1..=b.len() {
            let cost = if a[i - 1] == b[j - 1] { 0 } else { 1 };
            curr[j] = (prev[j] + 1)         // deletion
                .min(curr[j - 1] + 1)        // insertion
                .min(prev[j - 1] + cost);    // substitution
        }
        std::mem::swap(&mut prev, &mut curr);
    }
    prev[b.len()]
}

// ----- UI Automation: read the selected text in the focused element ----------

/// Tries to return the currently selected text in the focused element via
/// UI Automation's TextPattern. Works in most standard text controls.
unsafe fn get_selected_text(uia: &IUIAutomation) -> Option<String> {
    let focused = uia.GetFocusedElement().ok()?;
    let pattern = focused.GetCurrentPattern(UIA_TextPatternId).ok()?;
    let text_pattern = pattern.cast::<IUIAutomationTextPattern>().ok()?;
    let ranges = text_pattern.GetSelection().ok()?;
    if ranges.Length().ok()? == 0 {
        return None;
    }
    let range = ranges.GetElement(0).ok()?;
    let s = range.GetText(-1).ok()?.to_string();
    if s.trim().is_empty() { None } else { Some(s.trim().to_string()) }
}

/// Tries to get selected text by simulating Ctrl+C and reading the clipboard.
///
/// This works across a broad range of applications (including Adobe Acrobat)
/// that do not expose text selection through the UI Automation API.
///
/// Steps:
///   1. Save the current clipboard text (to restore it afterwards).
///   2. Clear the clipboard so we can detect a fresh copy.
///   3. Send Ctrl+C to the foreground window via SendInput.
///   4. Wait briefly for the application to process the keystroke.
///   5. Read the clipboard.  If it is empty the application had nothing
///      selected and we return None without disturbing the previous content.
///   6. Restore the previously saved clipboard text.
unsafe fn get_selected_via_clipboard() -> Option<String> {
    const CF_UNICODETEXT: u32 = 13;
    // Virtual key code for 'C' (no named constant in the windows crate).
    let vk_c = VIRTUAL_KEY(b'C' as u16);

    // ── 1. Save current clipboard text ──
    let saved: Option<Vec<u16>> = (|| -> Option<Vec<u16>> {
        OpenClipboard(HWND::default()).ok()?;
        let h = GetClipboardData(CF_UNICODETEXT).ok();
        let result = h.and_then(|h| {
            let ptr = GlobalLock(HGLOBAL(h.0)) as *const u16;
            if ptr.is_null() {
                return None;
            }
            let mut len = 0usize;
            while *ptr.add(len) != 0 {
                len += 1;
            }
            // Include the null terminator so we can restore exactly.
            let v = std::slice::from_raw_parts(ptr, len + 1).to_vec();
            let _ = GlobalUnlock(HGLOBAL(h.0));
            Some(v)
        });
        let _ = CloseClipboard();
        result
    })();

    // ── 2. Clear the clipboard ──
    if OpenClipboard(HWND::default()).is_ok() {
        let _ = EmptyClipboard();
        let _ = CloseClipboard();
    }

    // ── 3. Record the sequence number then send Ctrl+C ──
    let seq_before = GetClipboardSequenceNumber();

    let inputs: [INPUT; 4] = [
        INPUT {
            r#type: INPUT_KEYBOARD,
            Anonymous: INPUT_0 {
                ki: KEYBDINPUT {
                    wVk: VK_CONTROL,
                    wScan: 0,
                    dwFlags: KEYBD_EVENT_FLAGS(0),
                    time: 0,
                    dwExtraInfo: 0,
                },
            },
        },
        INPUT {
            r#type: INPUT_KEYBOARD,
            Anonymous: INPUT_0 {
                ki: KEYBDINPUT {
                    wVk: vk_c,
                    wScan: 0,
                    dwFlags: KEYBD_EVENT_FLAGS(0),
                    time: 0,
                    dwExtraInfo: 0,
                },
            },
        },
        INPUT {
            r#type: INPUT_KEYBOARD,
            Anonymous: INPUT_0 {
                ki: KEYBDINPUT {
                    wVk: vk_c,
                    wScan: 0,
                    dwFlags: KEYEVENTF_KEYUP,
                    time: 0,
                    dwExtraInfo: 0,
                },
            },
        },
        INPUT {
            r#type: INPUT_KEYBOARD,
            Anonymous: INPUT_0 {
                ki: KEYBDINPUT {
                    wVk: VK_CONTROL,
                    wScan: 0,
                    dwFlags: KEYEVENTF_KEYUP,
                    time: 0,
                    dwExtraInfo: 0,
                },
            },
        },
    ];
    SendInput(&inputs, std::mem::size_of::<INPUT>() as i32);

    // ── 4. Wait for the application to process the copy ──
    std::thread::sleep(std::time::Duration::from_millis(150));

    // ── 5. Read new clipboard content ──
    // If the sequence number has not changed, nothing was copied (no selection).
    let seq_after = GetClipboardSequenceNumber();
    let new_text: Option<String> = if seq_after == seq_before {
        None
    } else {
        (|| -> Option<String> {
            OpenClipboard(HWND::default()).ok()?;
            let h = GetClipboardData(CF_UNICODETEXT).ok()?;
            let ptr = GlobalLock(HGLOBAL(h.0)) as *const u16;
            if ptr.is_null() {
                let _ = CloseClipboard();
                return None;
            }
            let mut len = 0usize;
            while *ptr.add(len) != 0 {
                len += 1;
            }
            let s = String::from_utf16_lossy(std::slice::from_raw_parts(ptr, len));
            let _ = GlobalUnlock(HGLOBAL(h.0));
            let _ = CloseClipboard();
            let trimmed = s.trim().to_string();
            if trimmed.is_empty() { None } else { Some(trimmed) }
        })()
    };

    // ── 6. Restore previous clipboard content ──
    if OpenClipboard(HWND::default()).is_ok() {
        let _ = EmptyClipboard();
        if let Some(ref data) = saved {
            let byte_size = data.len() * std::mem::size_of::<u16>();
            if let Ok(hmem) = GlobalAlloc(GMEM_MOVEABLE, byte_size) {
                let ptr = GlobalLock(hmem) as *mut u16;
                if !ptr.is_null() {
                    std::ptr::copy_nonoverlapping(data.as_ptr(), ptr, data.len());
                    let _ = GlobalUnlock(hmem);
                    // SetClipboardData takes ownership of hmem on success; only
                    // free it ourselves if the call fails.
                    if SetClipboardData(CF_UNICODETEXT, HANDLE(hmem.0)).is_err() {
                        let _ = GlobalFree(hmem);
                    }
                } else {
                    let _ = GlobalFree(hmem);
                }
            }
        }
        let _ = CloseClipboard();
    }

    new_text
}

/// Returns the currently selected text, trying UI Automation first and
/// falling back to the clipboard simulation approach (Ctrl+C) for
/// applications that do not expose selection via the accessibility API.
unsafe fn get_text_at_cursor() -> Option<String> {
    let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);

    let uia: IUIAutomation =
        CoCreateInstance(&CUIAutomation, None, CLSCTX_INPROC_SERVER).ok()?;

    // ── 1. Try UI Automation selection ──────────────────────────────────────
    if let Some(s) = get_selected_text(&uia) {
        return Some(s);
    }

    // ── 2. Fall back to clipboard simulation (Ctrl+C) ───────────────────────
    // This handles applications such as Adobe Acrobat whose PDF rendering
    // layer does not expose a UIA TextPattern but does respond to Ctrl+C.
    get_selected_via_clipboard()
}

// ----- pick the most acronym-shaped token from whatever UIA returned -------
fn extract_candidate_acronym(text: &str) -> Option<String> {
    let tokens: Vec<&str> = text
        .split(|c: char| !c.is_ascii_alphanumeric() && c != '&')
        .filter(|s| !s.is_empty())
        .collect();

    let acronym = tokens
        .iter()
        .copied()
        .filter(|t| {
            t.len() >= 2
                && t.chars().any(|c| c.is_ascii_alphabetic())
                && t.chars().all(|c| c.is_ascii_uppercase() || c.is_ascii_digit() || c == '&')
        })
        .max_by_key(|s| s.len());

    acronym
        .map(str::to_string)
        .or_else(|| tokens.first().map(|s| s.to_string()))
}

// ----- formatting -----------------------------------------------------------
fn format_lookup(_term: &str, result: &LookupResult) -> (String, String) {
    match result {
        LookupResult::Exact(key, defs) => {
            let body = if defs.len() == 1 {
                defs[0].clone()
            } else {
                let mut s = format!("{} definitions found:\r\n\r\n", defs.len());
                for (i, d) in defs.iter().enumerate() {
                    s.push_str(&format!("{}. {}\r\n\r\n", i + 1, d));
                }
                s.trim_end().to_string()
            };
            (key.clone(), body)
        }
        LookupResult::Suggestions(query, hits) => {
            let mut s = format!(
                "'{}' is not in your acronyms file.\r\n\r\nDid you mean:\r\n\r\n",
                query
            );
            for (k, _d, defs) in hits {
                if defs.len() == 1 {
                    s.push_str(&format!("  {}  —  {}\r\n", k, defs[0]));
                } else {
                    s.push_str(&format!("  {}\r\n", k));
                    for d in defs {
                        s.push_str(&format!("      • {}\r\n", d));
                    }
                }
            }
            (format!("Not found: {}", query), s.trim_end().to_string())
        }
        LookupResult::NotFound(query, total) => (
            format!("Not found: {}", query),
            format!(
                "'{}' is not in your acronyms file, and no near matches were found.\r\n\r\n\
                 Searched {} definition(s).\r\n\r\n\
                 Press Ctrl+Shift+E to open the CSV and add it.",
                query, total
            ),
        ),
        LookupResult::Error(msg) => (
            "Error".to_string(),
            msg.replace('\n', "\r\n"),
        ),
    }
}

// ----- popup window (raw Win32) --------------------------------------------
//
// Runs on its own thread. The message loop drives the popup until the user
// closes it; the function returns and the thread exits. The popup_busy flag
// (set by the caller) ensures only one popup is alive at a time.
//
// Using raw Win32 avoids the native-windows-gui dependency on Shcore.dll
// (SetProcessDpiAwareness, Windows 8.1+) that caused STATUS_ENTRYPOINT_NOT_FOUND
// on older Windows versions.

/// Window procedure for the popup. Handles button click, Esc (sent by
/// IsDialogMessageW as WM_COMMAND(IDCANCEL=2)), and the × close button.
unsafe extern "system" fn popup_wnd_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    const IDCANCEL: usize = 2;
    match msg {
        WM_COMMAND => {
            // Low word of wParam is the control/notification ID.
            if wparam.0 & 0xffff == IDCANCEL {
                let _ = DestroyWindow(hwnd);
            }
            LRESULT(0)
        }
        WM_CLOSE => {
            let _ = DestroyWindow(hwnd);
            LRESULT(0)
        }
        WM_DESTROY => {
            PostQuitMessage(0);
            LRESULT(0)
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

fn show_popup(term: &str, result: LookupResult) {
    let (title, body) = format_lookup(term, &result);

    unsafe {
        let hmodule  = GetModuleHandleW(None).unwrap_or_default();
        let hinstance = HINSTANCE(hmodule.0);

        // Register the window class once per process; ignore "already exists".
        let class_name = windows::core::w!("AcronymLookupPopup");
        let wc = WNDCLASSEXW {
            cbSize:        std::mem::size_of::<WNDCLASSEXW>() as u32,
            lpfnWndProc:   Some(popup_wnd_proc),
            hInstance:     hinstance,
            hCursor:       LoadCursorW(None, IDC_ARROW).unwrap_or_default(),
            // Standard Win32 trick: passing (COLOR_BTNFACE+1) as a pseudo-brush.
            // COLOR_BTNFACE = 15, so 15+1 = 16.
            hbrBackground: HBRUSH(16 as *mut _),
            lpszClassName: class_name,
            ..Default::default()
        };
        let _ = RegisterClassExW(&wc);

        // Compute total window size from the desired 520×320 client area.
        let client_w: i32 = 520;
        let client_h: i32 = 320;
        let mut rect = RECT { left: 0, top: 0, right: client_w, bottom: client_h };
        let _ = AdjustWindowRectEx(
            &mut rect,
            WS_CAPTION | WS_SYSMENU,
            BOOL(0),
            WS_EX_TOPMOST,
        );
        let win_w = rect.right  - rect.left;
        let win_h = rect.bottom - rect.top;

        // Position near the cursor.
        let mut pt = POINT::default();
        let _ = GetCursorPos(&mut pt);

        let title_w: Vec<u16> = format!("Acronym - {}", title)
            .encode_utf16()
            .chain(std::iter::once(0))
            .collect();

        let hwnd = match CreateWindowExW(
            WS_EX_TOPMOST,
            class_name,
            PCWSTR(title_w.as_ptr()),
            WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
            pt.x + 16,
            pt.y + 16,
            win_w,
            win_h,
            HWND::default(),
            HMENU::default(),
            hinstance,
            None,
        ) {
            Ok(h) => h,
            Err(_) => return,
        };

        // Use the system GUI font (Segoe UI on Vista+, MS Shell Dlg 2 on XP).
        let hfont = GetStockObject(DEFAULT_GUI_FONT);

        // Scrollable read-only edit control covering most of the client area.
        let body_w: Vec<u16> = body.encode_utf16().chain(std::iter::once(0)).collect();
        if let Ok(hedit) = CreateWindowExW(
            WS_EX_CLIENTEDGE,
            windows::core::w!("EDIT"),
            PCWSTR(body_w.as_ptr()),
            WS_CHILD | WS_VISIBLE | WS_VSCROLL
                | WINDOW_STYLE((ES_MULTILINE | ES_READONLY | ES_AUTOVSCROLL) as u32),
            10,
            10,
            client_w - 24,
            client_h - 70,
            hwnd,
            HMENU::default(),
            hinstance,
            None,
        ) {
            SendMessageW(hedit, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
        }

        // Close button. ID = IDCANCEL (2) so IsDialogMessageW maps Esc → WM_COMMAND.
        if let Ok(hbtn) = CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            windows::core::w!("BUTTON"),
            windows::core::w!("Close (Esc)"),
            WS_CHILD | WS_VISIBLE,
            client_w - 124,
            client_h - 48,
            110,
            28,
            hwnd,
            HMENU(2 as *mut _),
            hinstance,
            None,
        ) {
            SendMessageW(hbtn, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
        }

        // Message loop. IsDialogMessageW translates Esc → WM_COMMAND(IDCANCEL=2)
        // and handles Tab focus navigation between the edit and button controls.
        let mut msg_buf = MSG::default();
        loop {
            let r = GetMessageW(&mut msg_buf, HWND::default(), 0, 0);
            if r.0 <= 0 {
                break;
            }
            if IsDialogMessageW(hwnd, &msg_buf).0 == 0 {
                let _ = TranslateMessage(&msg_buf);
                DispatchMessageW(&msg_buf);
            }
        }
    }
}

// ----- a tiny error box for the edit hotkey --------------------------------
fn error_box(message: &str) {
    let title: Vec<u16> = "Acronym Lookup\0".encode_utf16().collect();
    let body:  Vec<u16> = message.encode_utf16().chain(std::iter::once(0)).collect();
    unsafe {
        MessageBoxW(
            HWND::default(),
            PCWSTR(body.as_ptr()),
            PCWSTR(title.as_ptr()),
            MB_OK | MB_ICONERROR | MB_TOPMOST,
        );
    }
}
