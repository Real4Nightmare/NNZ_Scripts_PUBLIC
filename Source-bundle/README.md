# Source-bundle

Turn a whole project folder into **one readable text file** - every file path followed by its content, one after another - and rebuild the exact folder structure from that single file later.

Single script, no dependencies. Works on Windows PowerShell 5.1 and PowerShell 7+ (Windows, Linux, macOS).

```powershell
.\Source-bundle.ps1                        # interactive menu
.\Source-bundle.ps1 pack                   # project -> source-bundle.txt
.\Source-bundle.ps1 unpack -Target .\out   # source-bundle.txt -> project again
```

## What the bundle looks like

```text
=== SOURCE-BUNDLE v1 ===
=== CREATED: 2026-11-09T18:30:00.0000000Z ===
=== NAME: my-project ===
=== FILES: 3 ===
=== DIRS: 1 ===
=== COMPRESS: none ===
=== END-HEADER ===

=== INDEX: 3 file(s) ===
  docs/readme.md  (12 B)
  src/main.go  (29 B)
=== END-INDEX ===

@@@FILE-BEGIN:src/main.go
@@@META:enc=utf8;size=29;mtime=1760000000;sha256=...;eol=lf;nonl=0
package main

func main() {}
@@@FILE-END:src/main.go

@@@FILE-BEGIN:docs/readme.md
@@@META:enc=utf8;size=12;mtime=1760000000;sha256=...;eol=crlf;nonl=0
# Title
...
@@@FILE-END:docs/readme.md

@@@DIR:empty/folder

=== END OF BUNDLE ===
```

Everything is plain text you can read, search, paste into a chat, or commit. A file index at the top (optional, `-NoIndex` to drop it) shows what is inside before the contents start. Run `.\Source-bundle.ps1 schema` for the full format.

## Faithful restore

`unpack` rebuilds the project **byte for byte**:

- folder structure, including **empty folders**
- text files stay readable; binary files are stored as base64
- **line endings** (LF / CRLF), a **missing final newline**, a UTF-8 **BOM** and empty files are preserved; files with mixed line endings fall back to base64 so nothing changes
- every file carries a **SHA-256**; `verify` checks the whole bundle, and the packer double-checks each file's stored form before writing it
- a content line that looks like a marker (`@@@FILE-END:...`) is escaped, so bundles can contain other bundles
- restore refuses paths that would escape the target folder (`../..`, absolute paths)
- if Git's `autocrlf` converted the bundle to CRLF, it still restores correctly

## You choose what to do with `.gitignore`

Every pack asks (menu) or takes an option (command line):

| Choice | Option | Meaning |
|---|---|---|
| Apply ignore files | `-GitIgnore auto` (default) | built-in parser: root + nested ignore files, `.git/info/exclude` |
| Ask real git | `-GitIgnore git` | `git ls-files --cached --others --exclude-standard` |
| Ignore nothing | `-GitIgnore off` | everything is included |

More control:

- `-IgnoreFiles .gitignore,.dockerignore,.npmignore` - **which** ignore files are honoured (same syntax; `.bundleignore` is always supported for exclusions that only this tool should know about)
- `-SkipIgnoreFiles` - keep `.gitignore`, `.gitattributes`, ... themselves **out** of the bundle (default: they are packed)
- `-NoDefaultExcludes` - also include `node_modules`, `dist`, `build`, `.vscode`, ... (`.git` is never packed); `-ForceInclude` does the same for just the names you list
- `-NoEmptyDirs` - do not record empty folders

### Include something even though it is ignored (or the other way round)

Ignore files and the built-in default excludes (`node_modules`, `dist`, `build`, `.vscode`, ... - `.git` is never packed) are **separate switches**, so you can mix them:

| You want | Use |
|---|---|
| everything, literally | `-IncludeAll` (= `-GitIgnore off -NoDefaultExcludes`) |
| ignore files applied, but **all of `node_modules`** anyway | `-ForceInclude node_modules` |
| one ignored file anyway (e.g. `.env`) | `-ForceInclude .env` |
| a path or pattern anyway | `-ForceInclude 'dist/*','packages/a/node_modules'` |
| ignore files applied, a folder dropped anyway | `-ExcludeDir vendor` (always wins over `-ForceInclude`) |
| see what would go in before writing anything | `pack -DryRun` (file list + totals, nothing written) |

A plain name (`node_modules`) matches at any depth; a value with `/` or `*` is matched as a path pattern. Forced paths also skip the default file excludes (`*.dll`, `*.exe`, ...), so you really get the whole folder.

In the menu, choice **[4] "apply them, but let me pick folders/files to include anyway"** scans the project, lists what is being left out with sizes (e.g. `node_modules  12,345 files  310 MB  folder ignored by an ignore file`), and you pick by number. The choice is remembered and shown in the summary every time.

`gitcheck` shows what the ignore files remove (and warns about unignored `.env` / key files), `srctree` shows the live tree with the rules applied.

## Made for sharing projects with people and AI chats

| Need | How |
|---|---|
| Know how big it is for a chat | every pack prints a rough **token estimate** (`info` shows it too) |
| Bundle is too big for one message | `-SplitMB 5` writes `name.part1.txt`, `name.part2.txt`, ... - **each part is a complete bundle**; `unpack` them one after another into the same folder (or `merge` them) |
| Send only what you just changed | `-Changed` (uncommitted changes) or `-ChangedSince HEAD~3` / `main` - needs git |
| Send only what differs from the last bundle | `-Base old-bundle.txt` (incremental); unpack the old bundle, then the new one with `-Overwrite` |
| Don't leak secrets | `-ExcludeSecrets` leaves out `.env`, private keys, tokens, hard-coded passwords (it lists what it dropped) |
| Copy straight into a chat | `-Clipboard` copies the bundle text after packing |
| Rebuild a project someone pasted to you | copy the text, run `unpack -FromClipboard` (text wrapped in a Markdown code fence or surrounded by chat is fine) |

In the menu these are the commands `changed`, `incremental`, `split` and `paste`, plus options in the pack questions.

## Commands

| Command | What it does |
|---|---|
| *(none)* | interactive menu - pick **flat** (one list) or **grouped** (group, then command) |
| `pack` | folder -> bundle (`-Profile quick\|full\|code\|docs`, `-Exclude`, `-ExcludeDir`, `-Include`, `-OnlyExt`, `-NoBinary`, `-Since`, `-Compress`, plus the sharing options above) |
| `unpack` | bundle -> folder (`-Target`, `-Overwrite`, `-DryRun`, `-Filter '*.go'`, `-PreserveTimes`, `-FromClipboard`) |
| `extract` | restore one file or sub-folder (`-Path src/api`) |
| `sync` | update a folder from a bundle (`-DeleteExtras`; files protected by ignore files are never deleted) |
| `info` `list` `tree` `stats` | look inside (`-Long`, `-Depth`, `-OutputJson`) |
| `grep` | regex search in the bundle (`-Pattern`, `-Context`, `-CaseSensitive`) |
| `verify` | check every SHA-256 (exit code 1 on failure) |
| `diff` / `bdiff` | bundle vs folder / bundle vs bundle |
| `merge` | combine bundles (`-BundleFiles a.txt,b.txt -BundleFile out.txt`) |
| `duplicates` | files with identical content |
| `watch` | re-pack when files change (`-IntervalSec`) |
| `encrypt` / `decrypt` | AES-256-CBC + HMAC-SHA256, PBKDF2 (`-Key`), writes `<bundle>.enc` |
| `gitcheck` `srctree` `bigfiles` `secrets` `suggest` | ignore-file report, live tree, largest files, secret scan, suggested `.gitignore` lines |

`.\Source-bundle.ps1 help` prints all options.

The menu remembers your choices in `.source-bundle.json` (never the key). Typing a command name or a unique prefix works on every screen, `m` switches view, `r` repeats the last command.

## Notes and limits

- Text is stored as UTF-8. Files that are not valid UTF-8 are stored as base64.
- Defaults: text files up to 10 MB (`-MaxTextKB`), binary files up to 50 MB (`-MaxBinaryMB`); bigger files are listed as *skipped* and not stored.
- Symbolic links are not followed.
- A bundle is read into memory as a whole - fine for projects, not meant for multi-gigabyte trees.
- `secrets` is a heuristic scan, not a guarantee. Review a bundle before sharing it.
- `-Key` on the command line can end up in shell history; the menu asks for it without echo. A lost key cannot be recovered.
- Tested with PowerShell 7 on Linux. Windows PowerShell 5.1 uses the same code paths but was not run in testing.

---

## خلاصه (فارسی)

این اسکریپت کل یک پروژه را در **یک فایل متنی خوانا** قرار می‌دهد (مسیر هر فایل و بلافاصله محتوای آن، پشت سر هم) و بعداً از همان یک فایل کل ساختار پوشه‌ها و فایل‌ها را **دقیقاً مثل اول** بازسازی می‌کند، حتی پوشه‌های خالی، CRLF، فایل‌های باینری و فایل‌های بدون خط آخر.

- اجرای بدون آرگومان: منوی تعاملی (حالت **مسطح** یا **دسته‌بندی‌شده**).
- برای `.gitignore` و فایل‌های مشابه (`.dockerignore`، `.npmignore`، ...) خودت تصمیم می‌گیری: اعمال شوند، با خود git اعمال شوند، یا اصلاً اعمال نشوند. حتی می‌شود خود فایل‌های ایگنور را از بسته بیرون گذاشت.
- `pack` برای ساخت بسته و `unpack` برای بازسازی پروژه؛ `verify` درستی همه‌ی فایل‌ها را با SHA-256 بررسی می‌کند.
- بالای هر بسته یک **فهرست فایل‌ها** هست و تخمین **توکن** هم چاپ می‌شود.
- برای اینکه مثلاً `node_modules` با اینکه داخل `.gitignore` است وارد بسته شود: `-ForceInclude node_modules` یا در منو گزینه‌ی ۴ (فهرست چیزهای کنار گذاشته‌شده را با حجم نشان می‌دهد و با شماره انتخاب می‌کنی). `-IncludeAll` یعنی واقعاً همه چیز و `pack -DryRun` فقط پیش‌نمایش می‌دهد.
- برای ارسال به چت: `-SplitMB` (چند تکه، هر تکه یک بسته‌ی کامل)، `-Changed` (فقط فایل‌های تغییرکرده‌ی git)، `-Base` (افزایشی)، `-ExcludeSecrets` (حذف فایل‌های حساس)، `-Clipboard` (کپی به clipboard) و `unpack -FromClipboard` (ساخت پروژه از متن کپی‌شده).
