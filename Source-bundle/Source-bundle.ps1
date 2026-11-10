<#
.SYNOPSIS
  Source-bundle - turn a whole project folder into ONE readable text file, and rebuild it again.

.DESCRIPTION
  pack     every file path + its content, one after another, in a single text file
  unpack   rebuild the exact folder structure, files and bytes from that one file
  Also: list / tree / info / stats / grep / verify / diff / merge / sync / watch /
        encrypt / decrypt / duplicates and a few .gitignore-aware helper tools.

  Text files are stored as plain readable text. Binary files are stored as base64.
  Line endings (LF / CRLF), missing final newlines, empty files and empty folders
  are all preserved, and every file carries a SHA-256 so a restore can be verified.

  What to do with .gitignore (and similar ignore files) is a choice you make:
    -GitIgnore auto   apply ignore files with the built-in parser (default)
    -GitIgnore git    ask real git (git ls-files)
    -GitIgnore off    ignore nothing, include everything
  -IgnoreFiles picks WHICH ignore files are honoured (default .gitignore,.bundleignore;
  add .dockerignore, .npmignore ... if you want). -SkipIgnoreFiles keeps the ignore
  files themselves out of the bundle.

.EXAMPLE
  .\Source-bundle.ps1                       # interactive menu (flat or grouped)
  .\Source-bundle.ps1 pack                  # whole project -> source-bundle.txt
  .\Source-bundle.ps1 pack -GitIgnore off   # include everything, ignore nothing
  .\Source-bundle.ps1 unpack -Target .\out  # rebuild the project from the bundle
  .\Source-bundle.ps1 tree                  # show what is inside
#>
param(
    [Parameter(Position = 0)]
    [ValidateSet('interactive','pack','unpack','info','list','tree','extract','grep','verify','stats','diff','merge','sync','watch','encrypt','decrypt','duplicates','gitcheck','srctree','bigfiles','secrets','suggest','bdiff','schema','help')]
    [string]$Mode = 'interactive',

    # interactive view: ask first (default), or force flat / grouped
    [ValidateSet('ask','flat','grouped')]
    [string]$Menu = 'ask',

    [string]$BundleFile = 'source-bundle.txt',
    [string]$Source = '.',
    [string]$Target = 'restored',

    # ---- what goes into the bundle ----
    [ValidateSet('','quick','full','code','docs')]
    [string]$Profile = '',
    [string[]]$Exclude = @(),
    [string[]]$ExcludeDir = @(),
    [string[]]$Include = @(),
    [string[]]$OnlyExt = @(),
    [switch]$NoBinary,
    [int]$MaxBinaryMB = 50,
    [int]$MaxTextKB = 10240,
    [datetime]$Since = [datetime]::MinValue,
    [switch]$Compress,
    [double]$SplitMB = 0,           # split the bundle into parts of about N MB (each part is a complete bundle)
    [switch]$NoIndex,               # do not write the file index at the top of the bundle
    [switch]$Clipboard,             # pack: also copy the bundle text to the clipboard
    [switch]$ExcludeSecrets,        # pack: leave out files that look like they contain secrets
    [switch]$Changed,               # pack: only files git reports as changed (working tree)
    [string]$ChangedSince = '',     # pack: only files changed since this git ref (e.g. HEAD~3, main)
    [string]$Base = '',             # pack: incremental - only files that differ from this earlier bundle
    [switch]$NoDefaultExcludes,     # also include node_modules, dist, build ... (.git is always skipped)
    [switch]$NoEmptyDirs,           # do not record empty folders

    # ---- ignore files (.gitignore & friends) ----
    [ValidateSet('auto','git','off')]
    [string]$GitIgnore = 'auto',
    [string[]]$IgnoreFiles = @('.gitignore','.bundleignore'),
    [string[]]$ForceInclude = @(),  # include these even if an ignore file or the default excludes would drop them (e.g. node_modules)
    [switch]$IncludeAll,            # shortcut: -GitIgnore off + -NoDefaultExcludes (literally everything except .git)
    [switch]$SkipIgnoreFiles,

    # ---- unpack / sync / extract ----
    [switch]$FromClipboard,         # read the bundle from the clipboard instead of a file
    [switch]$Overwrite,
    [switch]$DryRun,
    [string]$Filter = '*',
    [string]$Path = '',
    [switch]$DeleteExtras,
    [switch]$PreserveTimes,

    # ---- inspect / search ----
    [string]$Pattern = '',
    [switch]$CaseSensitive,
    [int]$Context = 0,
    [switch]$LineNumbers,
    [switch]$Long,
    [int]$Depth = 0,
    [string]$CompareTo = '.',
    [string]$Compare2 = '',
    [string[]]$BundleFiles = @(),
    [int]$Top = 20,
    [switch]$OutputJson,

    # ---- encryption / watch ----
    [string]$Key = '',
    [int]$IntervalSec = 3
)

$ErrorActionPreference = 'Stop'

# `powershell -File x.ps1 -Exclude a,b` delivers "a,b" as ONE string - split list options on commas
$Exclude      = @($Exclude      | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
$ExcludeDir   = @($ExcludeDir   | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
$Include      = @($Include      | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
$OnlyExt      = @($OnlyExt      | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim().TrimStart('*').TrimStart('.').ToLowerInvariant() } | Where-Object { $_ })
$IgnoreFiles  = @($IgnoreFiles  | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
$ForceInclude  = @($ForceInclude  | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim().Trim('/','\') } | Where-Object { $_ })
$BundleFiles  = @($BundleFiles  | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
try { Add-Type -AssemblyName System.IO.Compression } catch {}
$utf8 = New-Object System.Text.UTF8Encoding($false)
$script:AppVersion = '1.0.0'
$script:FormatName = 'SOURCE-BUNDLE'
$script:VerifyFailed = $false
$sep = [System.IO.Path]::DirectorySeparatorChar

$alwaysExcludeDirs  = @('.git', '.svn', '.hg')
$defaultExcludeDirs = @('node_modules','__pycache__','.venv','dist','build','target','.next','.nuxt','.idea','.vscode','.pytest_cache','.cache','coverage','.vs','.gradle','.mvn','.terraform','.tox','.mypy_cache')
$defaultExcludeFiles = @('*.pyc','*.pyo','*.exe','*.dll','*.so','*.dylib','*.o','*.a','*.obj','*.lib','*.class','*.jar','*.war','*.pdb','*.ilk','*.exp','*.swp','*.swo','*~')
$alwaysExcludeFiles = @('.source-bundle.json','*.sbtmp')
$script:SkipDirs = @()
if ($IncludeAll) { $GitIgnore = 'off'; $NoDefaultExcludes = [switch]$true }

function Initialize-Excludes {
    $d = @($alwaysExcludeDirs)
    if (-not $NoDefaultExcludes) { $d += $defaultExcludeDirs }
    $script:SkipDirs = $d
}

# ============================================================
# help / format spec
# ============================================================
function Write-Help {
    Write-Host "Source-bundle v$script:AppVersion - project folder <-> one readable text file" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Usage: .\Source-bundle.ps1 [command] [options]" -ForegroundColor Yellow
    Write-Host ""
    $cmdText = @'
(none)|Interactive menu: choose FLAT or GROUPED view
pack|Folder -> one text file (paths + contents, one after another)
unpack|One text file -> the original folder structure and files
info|Bundle header and counts
list|List files (-Long for hash/mtime/line endings)
tree|Show the bundle as a tree (-Depth N)
stats|Sizes by extension and top folder
extract|Restore one file or sub-folder (-Path src/api)
grep|Regex search inside the bundle (-Pattern, -Context)
verify|Check SHA-256 of every file in the bundle
diff|Bundle vs a live folder (-CompareTo)
bdiff|Bundle vs bundle (-BundleFile, -Compare2)
merge|Combine bundles (-BundleFiles a,b -BundleFile out)
sync|Update a folder from the bundle (-DeleteExtras optional)
watch|Re-pack automatically when files change
encrypt|AES-256 encrypt a bundle with -Key (writes <bundle>.enc)
decrypt|Decrypt a .enc bundle with -Key
duplicates|Find files with identical content
gitcheck|What the ignore files exclude + risky files
srctree|Live folder tree with the ignore rules applied
bigfiles|Largest files that would be packed (-Top N)
secrets|Heuristic scan for keys/tokens/passwords
suggest|Suggest missing .gitignore patterns (-Overwrite appends)
schema|Print the bundle format
'@
    foreach ($line in ($cmdText -split "`r?`n")) { if ($line) { $p = $line.Split('|', 2); Write-Host ("  {0,-12} {1}" -f $p[0],$p[1]) } }
    Write-Host ""
    Write-Host "What to pack:" -ForegroundColor Yellow
    Write-Host "  -Profile quick|full|code|docs   -Exclude '*.log'   -ExcludeDir vendor   -OnlyExt go,md"
    Write-Host "  -NoBinary   -Compress (gzip)   -NoDefaultExcludes (include node_modules, dist, ...)   -NoEmptyDirs"
    Write-Host "  -ForceInclude node_modules,dist   include these even if an ignore file / the default excludes drop them"
    Write-Host "  -IncludeAll           literally everything (= -GitIgnore off -NoDefaultExcludes)"
    Write-Host "  -DryRun (with pack)   preview the file list, write nothing"
    Write-Host "  -SplitMB 5            several parts of ~5 MB, each a complete bundle (chat size limits)"
    Write-Host "  -Changed | -ChangedSince HEAD~3   only files git reports as changed"
    Write-Host "  -Base old-bundle.txt  incremental: only files that differ from an earlier bundle"
    Write-Host "  -ExcludeSecrets       leave out .env / keys / tokens     -Clipboard   copy the bundle text too"
    Write-Host "  -NoIndex              no file index at the top of the bundle"
    Write-Host "  unpack -FromClipboard rebuild a project from bundle text on the clipboard"
    Write-Host ""
    Write-Host "Ignore files (your choice):" -ForegroundColor Yellow
    Write-Host "  -GitIgnore auto|git|off     apply ignore files / ask real git / ignore nothing"
    Write-Host "  -IgnoreFiles .gitignore,.dockerignore   which ignore files to honour (+ .bundleignore)"
    Write-Host "  -SkipIgnoreFiles            leave the ignore files themselves out of the bundle"
    Write-Host ""
    Write-Host "Examples:" -ForegroundColor Yellow
    Write-Host "  .\Source-bundle.ps1 pack"
    Write-Host "  .\Source-bundle.ps1 pack -GitIgnore off -NoDefaultExcludes   # literally everything"
    Write-Host "  .\Source-bundle.ps1 unpack -Target .\rebuilt"
    Write-Host "  .\Source-bundle.ps1 grep -Pattern 'TODO' -Context 2"
}

function Write-Schema {
    Write-Host @"
Source-bundle format v1  (plain UTF-8 text, LF line endings)

=== SOURCE-BUNDLE v1 ===
=== CREATED: <ISO-8601 UTC> ===
=== NAME: <source folder name> ===
=== FILES: <count> ===
=== DIRS: <count of empty folders> ===
=== COMPRESS: <none|gzip> ===
[=== BASE: <earlier bundle> ===]        (incremental bundles only)
=== END-HEADER ===

=== INDEX: <count> file(s) ===          (optional table of contents, skipped by the reader)
  <relative/path>  (<size>)
=== END-INDEX ===

@@@FILE-BEGIN:<relative/path>
@@@META:enc=<utf8|base64|empty|skipped>;size=<bytes>;mtime=<unix>;sha256=<hex>[;eol=<lf|crlf>;nonl=<0|1>][;reason=<text>]
<content: the file text as-is, or base64 lines>
@@@FILE-END:<relative/path>

@@@DIR:<relative/empty/folder>

=== END OF BUNDLE ===

Rules:
  * utf8 text keeps its real line endings (eol) and missing final newline (nonl).
  * A content line that starts with @@@ is stored with one extra @ (unescaped on read).
  * binary files, and text with mixed line endings, are stored as base64.
  * skipped = listed but not stored (too large / binary skipped); restore ignores it.
  * Encrypted bundles start with === SOURCE-BUNDLE-ENCRYPTED v1 === (AES-256-CBC + HMAC-SHA256).
"@
}

# ============================================================
# helpers
# ============================================================
function Test-BytesBinary([byte[]]$b) {
    if ($b.Length -eq 0) { return $false }
    $n = [Math]::Min(8191, $b.Length - 1)
    for ($i = 0; $i -le $n; $i++) { if ($b[$i] -eq 0) { return $true } }
    try { [void](New-Object System.Text.UTF8Encoding($false,$true)).GetString($b); return $false } catch { return $true }
}
function Test-IsBinary([string]$p) { return (Test-BytesBinary ([System.IO.File]::ReadAllBytes($p))) }
function Get-Rel([string]$root,[string]$full) { $full.Substring($root.Length).TrimStart('\','/').Replace('\','/') }
function Get-Sha256([byte[]]$bytes) {
    $h = [System.Security.Cryptography.SHA256]::Create()
    try { return [System.BitConverter]::ToString($h.ComputeHash($bytes)).Replace('-','').ToLowerInvariant() } finally { $h.Dispose() }
}
function Get-Sha256File([string]$p) { Get-Sha256 ([System.IO.File]::ReadAllBytes($p)) }
function Format-Size([long]$n) {
    if ($n -lt 1024) { "$n B" }
    elseif ($n -lt 1048576) { "$([math]::Round($n/1KB,1)) KB" }
    elseif ($n -lt 1073741824) { "$([math]::Round($n/1MB,1)) MB" }
    else { "$([math]::Round($n/1GB,2)) GB" }
}
# a path stored in a bundle must never escape the folder we restore into
function Test-SafeRel([string]$rel) {
    if ([string]::IsNullOrWhiteSpace($rel)) { return $false }
    if ($rel.Contains([string][char]0)) { return $false }
    if ($rel.StartsWith('/') -or $rel.StartsWith('\') -or $rel -match '^[A-Za-z]:') { return $false }
    foreach ($seg in ($rel -split '[/\\]')) { if ($seg -eq '..') { return $false } }
    return $true
}
function Join-Safe([string]$root,[string]$rel) {
    if (-not (Test-SafeRel $rel)) { return $null }
    $rootFull = [System.IO.Path]::GetFullPath($root).TrimEnd($sep)
    $full = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($rootFull, $rel.Replace('/', $sep)))
    if ($full -eq $rootFull -or $full.StartsWith($rootFull + $sep, [System.StringComparison]::Ordinal)) { return $full }
    return $null
}
function Get-AbsPath([string]$p) {
    if ([System.IO.Path]::IsPathRooted($p)) { return [System.IO.Path]::GetFullPath($p) }
    return [System.IO.Path]::GetFullPath((Join-Path (Get-Location).Path $p))
}

# ============================================================
# encryption: AES-256-CBC + HMAC-SHA256 (encrypt-then-MAC), PBKDF2 key derivation
# ============================================================
$script:KdfIter = 210000
function Get-KeyMaterial([string]$pass, [byte[]]$salt) {
    $kdf = New-Object System.Security.Cryptography.Rfc2898DeriveBytes($pass, $salt, $script:KdfIter)
    try { $km = $kdf.GetBytes(64) } finally { $kdf.Dispose() }
    $ek = New-Object byte[] 32; $mk = New-Object byte[] 32
    [Array]::Copy($km, 0, $ek, 0, 32); [Array]::Copy($km, 32, $mk, 0, 32)
    return @{ Enc = $ek; Mac = $mk }
}
function Protect-Bytes([byte[]]$plain, [string]$pass) {
    $salt = New-Object byte[] 16; $iv = New-Object byte[] 16
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create(); $rng.GetBytes($salt); $rng.GetBytes($iv); $rng.Dispose()
    $k = Get-KeyMaterial $pass $salt
    $aes = [System.Security.Cryptography.Aes]::Create()
    $aes.KeySize = 256; $aes.Mode = [System.Security.Cryptography.CipherMode]::CBC; $aes.Padding = [System.Security.Cryptography.PaddingMode]::PKCS7
    $aes.Key = $k.Enc; $aes.IV = $iv
    $ct = $aes.CreateEncryptor().TransformFinalBlock($plain, 0, $plain.Length); $aes.Dispose()
    $ms = New-Object System.IO.MemoryStream
    $ms.Write($salt,0,16); $ms.Write($iv,0,16); $ms.Write($ct,0,$ct.Length)
    $body = $ms.ToArray(); $ms.Dispose()
    $hm = New-Object System.Security.Cryptography.HMACSHA256(,$k.Mac)
    $tag = $hm.ComputeHash($body); $hm.Dispose()
    $out = New-Object byte[] ($body.Length + 32)
    [Array]::Copy($body,0,$out,0,$body.Length); [Array]::Copy($tag,0,$out,$body.Length,32)
    return $out
}
function Unprotect-Bytes([byte[]]$blob, [string]$pass) {
    if ($blob.Length -lt 16 + 16 + 16 + 32) { throw "encrypted data is too short / corrupted" }
    $salt = New-Object byte[] 16; $iv = New-Object byte[] 16
    [Array]::Copy($blob,0,$salt,0,16); [Array]::Copy($blob,16,$iv,0,16)
    $bodyLen = $blob.Length - 32
    $body = New-Object byte[] $bodyLen; [Array]::Copy($blob,0,$body,0,$bodyLen)
    $tag = New-Object byte[] 32; [Array]::Copy($blob,$bodyLen,$tag,0,32)
    $k = Get-KeyMaterial $pass $salt
    $hm = New-Object System.Security.Cryptography.HMACSHA256(,$k.Mac)
    $calc = $hm.ComputeHash($body); $hm.Dispose()
    $diff = 0; for ($i = 0; $i -lt 32; $i++) { $diff = $diff -bor ($calc[$i] -bxor $tag[$i]) }
    if ($diff -ne 0) { throw "wrong key or corrupted bundle" }
    $ctLen = $bodyLen - 32
    $ct = New-Object byte[] $ctLen; [Array]::Copy($blob,32,$ct,0,$ctLen)
    $aes = [System.Security.Cryptography.Aes]::Create()
    $aes.KeySize = 256; $aes.Mode = [System.Security.Cryptography.CipherMode]::CBC; $aes.Padding = [System.Security.Cryptography.PaddingMode]::PKCS7
    $aes.Key = $k.Enc; $aes.IV = $iv
    try { return $aes.CreateDecryptor().TransformFinalBlock($ct, 0, $ct.Length) } finally { $aes.Dispose() }
}
$script:EncHeader = '=== SOURCE-BUNDLE-ENCRYPTED v1 ==='

# ============================================================
# bundle I/O
# ============================================================
function Get-EntryBytes($e) {
    switch ($e.Enc) {
        'empty'  { return ,[byte[]]@() }
        'base64' { return [Convert]::FromBase64String(($e.Content -replace '\s','')) }
        'utf8'   {
            $t = $e.Content
            if ([int]$e.NoNl -eq 1 -and $t.EndsWith("`n")) { $t = $t.Substring(0, $t.Length - 1) }
            if ($e.Eol -eq 'crlf') { $t = $t.Replace("`n", "`r`n") }
            return $utf8.GetBytes($t)
        }
    }
    return $null
}

function New-FileEntry([string]$full, [string]$rel) {
    $fi = New-Object System.IO.FileInfo $full
    $size = [long]$fi.Length
    $mtime = [long]([DateTimeOffset]$fi.LastWriteTimeUtc).ToUnixTimeSeconds()
    $e = [pscustomobject]@{ Path = $rel; Enc = 'utf8'; Size = $size; MTime = $mtime; Hash = ''; Eol = 'lf'; NoNl = 0; Reason = ''; Content = '' }
    if ($size -eq 0) { $e.Enc = 'empty'; $e.Hash = Get-Sha256 ([byte[]]@()); return $e }
    if ($size -gt ([long]$MaxBinaryMB * 1MB)) { $e.Enc = 'skipped'; $e.Reason = 'too-large'; return $e }
    $bytes = [System.IO.File]::ReadAllBytes($full)
    $e.Hash = Get-Sha256 $bytes
    if (Test-BytesBinary $bytes) {
        if ($NoBinary) { $e.Enc = 'skipped'; $e.Reason = 'binary-skipped'; $e.Hash = ''; return $e }
        $e.Enc = 'base64'; $e.Content = [Convert]::ToBase64String($bytes); return $e
    }
    if ($size -gt ([long]$MaxTextKB * 1KB)) { $e.Enc = 'skipped'; $e.Reason = 'text-too-large'; $e.Hash = ''; return $e }
    $text = $utf8.GetString($bytes)
    $crlf = [regex]::Matches($text, "`r`n").Count
    $lf   = [regex]::Matches($text, "(?<!`r)`n").Count
    $cr   = [regex]::Matches($text, "`r(?!`n)").Count
    if ($cr -gt 0 -or ($crlf -gt 0 -and $lf -gt 0)) { $e.Enc = 'base64'; $e.Content = [Convert]::ToBase64String($bytes); return $e }   # mixed endings: keep exact bytes
    $norm = if ($crlf -gt 0) { $text.Replace("`r`n", "`n") } else { $text }
    if ($crlf -gt 0) { $e.Eol = 'crlf' }
    if (-not $norm.EndsWith("`n")) { $e.NoNl = 1; $norm += "`n" }
    $e.Content = $norm
    # safety net: the stored form must rebuild the original bytes exactly, otherwise fall back to base64
    if ((Get-Sha256 (Get-EntryBytes $e)) -ne $e.Hash) { $e.Enc = 'base64'; $e.Eol = 'lf'; $e.NoNl = 0; $e.Content = [Convert]::ToBase64String($bytes) }
    return $e
}

function Write-BundleHeader($w, [string]$name, [int]$files, [int]$dirs, [bool]$gz, $index = $null, $dirList = $null, [string]$baseName = '') {
    $w.WriteLine("=== $script:FormatName v1 ===")
    $w.WriteLine("=== CREATED: $((Get-Date).ToUniversalTime().ToString('o')) ===")
    $w.WriteLine("=== NAME: $name ===")
    $w.WriteLine("=== FILES: $files ===")
    $w.WriteLine("=== DIRS: $dirs ===")
    $w.WriteLine("=== COMPRESS: $(if ($gz) { 'gzip' } else { 'none' }) ===")
    if ($baseName) { $w.WriteLine("=== BASE: $baseName ===") }
    $w.WriteLine("=== END-HEADER ===")
    # human-friendly table of contents; the reader skips it
    if ($index -and -not $NoIndex) {
        $w.WriteLine("")
        $w.WriteLine("=== INDEX: $(@($index).Count) file(s) ===")
        foreach ($it in $index) { $w.WriteLine("  $($it.Rel)  ($(Format-Size ([long]$it.Size)))") }
        foreach ($d in @($dirList)) { if ($d) { $w.WriteLine("  $d/  (empty folder)") } }
        $w.WriteLine("=== END-INDEX ===")
    }
    $w.WriteLine("")
}

function Write-BundleEntry($w, $e) {
    $w.WriteLine("@@@FILE-BEGIN:$($e.Path)")
    $meta = "enc=$($e.Enc);size=$($e.Size);mtime=$($e.MTime);sha256=$($e.Hash)"
    if ($e.Enc -eq 'utf8') { $meta += ";eol=$($e.Eol);nonl=$($e.NoNl)" }
    if ($e.Reason) { $meta += ";reason=$($e.Reason)" }
    $w.WriteLine("@@@META:$meta")
    switch ($e.Enc) {
        'utf8'    { $w.Write([regex]::Replace($e.Content, '(?m)^@@@', '@@@@')) }
        'base64'  { $b = $e.Content; for ($i = 0; $i -lt $b.Length; $i += 76) { $w.WriteLine($b.Substring($i, [Math]::Min(76, $b.Length - $i))) } }
        'skipped' { $w.WriteLine("[skipped: $($e.Reason)]") }
    }
    $w.WriteLine("@@@FILE-END:$($e.Path)")
    $w.WriteLine("")
}
function Write-BundleFooter($w, $dirs) {
    foreach ($d in ($dirs | Sort-Object)) { $w.WriteLine("@@@DIR:$d") }
    $w.WriteLine("")
    $w.WriteLine("=== END OF BUNDLE ===")
}
# finish: optionally gzip the temp file into its final place
function Complete-BundleFile([string]$tmp, [string]$dest, [bool]$gz) {
    if ($gz) {
        $inS = [System.IO.File]::OpenRead($tmp); $outS = [System.IO.File]::Create($dest)
        $z = New-Object System.IO.Compression.GZipStream($outS, [System.IO.Compression.CompressionLevel]::Optimal)
        try { $inS.CopyTo($z) } finally { $z.Dispose(); $outS.Dispose(); $inS.Dispose() }
        Remove-Item $tmp -Force
    } else { Move-Item $tmp $dest -Force }
}

function Get-BundleText([string]$path, [string]$pass) {
    $bytes = [System.IO.File]::ReadAllBytes($path)
    $head = [System.Text.Encoding]::ASCII.GetString($bytes, 0, [Math]::Min(40, $bytes.Length))
    if ($head.StartsWith('=== SOURCE-BUNDLE-ENCRYPTED')) {
        if (-not $pass) { throw "bundle is encrypted; supply -Key" }
        $lines = ($utf8.GetString($bytes) -replace "`r", '') -split "`n"
        $b64 = ($lines | Where-Object { $_ -and -not $_.StartsWith('===') }) -join ''
        $bytes = Unprotect-Bytes ([Convert]::FromBase64String($b64)) $pass
    }
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0x1F -and $bytes[1] -eq 0x8B) {
        $ms = New-Object System.IO.MemoryStream(,$bytes)
        $z = New-Object System.IO.Compression.GZipStream($ms, [System.IO.Compression.CompressionMode]::Decompress)
        $mem = New-Object System.IO.MemoryStream
        $z.CopyTo($mem); $z.Dispose(); $ms.Dispose()
        $bytes = $mem.ToArray(); $mem.Dispose()
    }
    $text = $utf8.GetString($bytes)
    if ($text.Length -gt 0 -and $text[0] -eq [char]0xFEFF) { $text = $text.Substring(1) }
    $nl = $text.IndexOf("`n")
    if ($nl -gt 0 -and $text[$nl - 1] -eq "`r") { $text = $text.Replace("`r`n", "`n") }   # file was converted to CRLF (e.g. git autocrlf)
    if (-not $text.StartsWith("=== $script:FormatName v")) { throw "not a Source-bundle file: $path" }
    return $text
}

$script:TempFiles = @()
function Clear-TempFiles { foreach ($t in $script:TempFiles) { Remove-Item -LiteralPath $t -Force -ErrorAction SilentlyContinue }; $script:TempFiles = @() }

# path of the bundle to read: the -BundleFile, or the clipboard text saved to a temp file
function Get-BundleInput {
    if (-not $FromClipboard) { return (Resolve-Path $BundleFile).Path }
    if (-not (Get-Command Get-Clipboard -ErrorAction SilentlyContinue)) { throw "Get-Clipboard is not available on this system" }
    $t = Get-Clipboard -Raw
    if ([string]::IsNullOrWhiteSpace($t)) { throw "the clipboard is empty" }
    $a = $t.IndexOf("=== $script:FormatName v")
    if ($a -lt 0) { $a = $t.IndexOf('=== SOURCE-BUNDLE-ENCRYPTED') }
    if ($a -lt 0) { throw "no Source-bundle text found on the clipboard" }
    $t = $t.Substring($a)                                   # drop a leading ``` fence or chatter
    $z = $t.LastIndexOf('=== END OF BUNDLE ===')
    if ($z -ge 0) { $t = $t.Substring(0, $z + '=== END OF BUNDLE ==='.Length) + "`n" }
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ("sb-clip-" + [guid]::NewGuid().ToString('N') + '.txt')
    [System.IO.File]::WriteAllText($tmp, $t, $utf8)
    $script:TempFiles += $tmp
    return $tmp
}

function Read-Bundle([string]$path, [string]$pass) {
    $text = Get-BundleText $path $pass
    $lines = $text.Split("`n")
    $h = @{ Version = 1; Created = ''; Name = ''; Files = 0; Dirs = 0; Compress = 'none'; Base = '' }
    $entries = New-Object System.Collections.Generic.List[object]
    $dirs = New-Object System.Collections.Generic.List[string]
    $cur = $null; $buf = New-Object System.Text.StringBuilder; $needMeta = $false; $inHeader = $true
    foreach ($line in $lines) {
        if ($inHeader) {
            if ($line -match '^=== (CREATED|NAME|FILES|DIRS|COMPRESS|BASE): (.*) ===$') { $h[$matches[1].Substring(0,1) + $matches[1].Substring(1).ToLowerInvariant()] = $matches[2]; continue }
            if ($line -eq '=== END-HEADER ===') { $inHeader = $false }
            continue
        }
        if ($null -eq $cur) {
            if ($line.StartsWith('@@@FILE-BEGIN:')) {
                $cur = @{ Path = $line.Substring(14); Enc = 'utf8'; Size = [long]0; MTime = [long]0; Hash = ''; Eol = 'lf'; NoNl = 0; Reason = ''; Content = '' }
                $buf.Clear() | Out-Null; $needMeta = $true
            } elseif ($line.StartsWith('@@@DIR:')) { $dirs.Add($line.Substring(7)) }
            continue
        }
        if ($needMeta) {
            $needMeta = $false
            if ($line.StartsWith('@@@META:')) {
                foreach ($kv in $line.Substring(8).Split(';')) {
                    $ix = $kv.IndexOf('='); if ($ix -lt 1) { continue }
                    $k = $kv.Substring(0,$ix); $v = $kv.Substring($ix + 1)
                    switch ($k) {
                        'enc'    { $cur.Enc = $v }
                        'size'   { $cur.Size = [long]$v }
                        'mtime'  { $cur.MTime = [long]$v }
                        'sha256' { $cur.Hash = $v }
                        'eol'    { $cur.Eol = $v }
                        'nonl'   { $cur.NoNl = [int]$v }
                        'reason' { $cur.Reason = $v }
                    }
                }
                continue
            }
        }
        if ($line.StartsWith('@@@FILE-END:')) {
            $cur.Content = $buf.ToString()
            $entries.Add([pscustomobject]$cur); $cur = $null
        } elseif ($line.StartsWith('@@@@')) { [void]$buf.Append($line.Substring(1)).Append("`n") }
        else { [void]$buf.Append($line).Append("`n") }
    }
    if ($null -ne $cur) { Write-Warning "bundle is truncated: last file '$($cur.Path)' is incomplete and was dropped" }
    return @{ Header = $h; Entries = $entries; Dirs = $dirs; Complete = $text.TrimEnd().EndsWith('=== END OF BUNDLE ===') }
}

# ============================================================
# ignore-file engine (.gitignore syntax)
#   -GitIgnore auto  built-in parser: the ignore files named in -IgnoreFiles (root + nested folders)
#                    and .git/info/exclude.  .bundleignore uses the same syntax and is meant for
#                    exclusions that only this tool should know about.
#   -GitIgnore git   ask real git: git ls-files --cached --others --exclude-standard
#   -GitIgnore off   ignore nothing
# ============================================================
function ConvertTo-GitRegex([string]$glob) {
    $sb = New-Object System.Text.StringBuilder
    $i = 0; $n = $glob.Length
    while ($i -lt $n) {
        $ch = $glob.Substring($i,1)
        switch ($ch) {
            '\' { if ($i + 1 -lt $n) { [void]$sb.Append([regex]::Escape($glob.Substring($i+1,1))); $i++ } }
            '*' {
                if ($i + 1 -lt $n -and $glob.Substring($i+1,1) -eq '*') {
                    $i++
                    if ($i + 1 -lt $n -and $glob.Substring($i+1,1) -eq '/') { [void]$sb.Append('(?:.*/)?'); $i++ }
                    else { [void]$sb.Append('.*') }
                } else { [void]$sb.Append('[^/]*') }
            }
            '?' { [void]$sb.Append('[^/]') }
            '[' {
                $j = if ($i + 2 -lt $n) { $glob.IndexOf(']', $i + 2) } else { -1 }
                if ($j -gt 0) {
                    $cls = $glob.Substring($i + 1, $j - $i - 1)
                    if ($cls.StartsWith('!')) { $cls = '^' + $cls.Substring(1) }
                    [void]$sb.Append('[' + $cls.Replace('\','\\') + ']'); $i = $j
                } else { [void]$sb.Append('\[') }
            }
            default { [void]$sb.Append([regex]::Escape($ch)) }
        }
        $i++
    }
    return $sb.ToString()
}

function Read-GitIgnoreFile([string]$file, [string]$baseDir, [string]$label, [bool]$icase) {
    $out = @()
    $opt = if ($icase) { [System.Text.RegularExpressions.RegexOptions]::IgnoreCase } else { [System.Text.RegularExpressions.RegexOptions]::None }
    foreach ($raw in [System.IO.File]::ReadAllLines($file)) {
        $line = [regex]::Replace($raw, '(?<!\\)\s+$', '')
        if ($line -eq '' -or $line.StartsWith('#')) { continue }
        $neg = $false
        if ($line.StartsWith('!')) { $neg = $true; $line = $line.Substring(1) }
        elseif ($line.StartsWith('\!') -or $line.StartsWith('\#')) { $line = $line.Substring(1) }
        $dirOnly = $false
        if ($line.EndsWith('/')) { $dirOnly = $true; $line = $line.TrimEnd('/') }
        if ($line -eq '') { continue }
        $anch = $line.Contains('/')
        if ($line.StartsWith('/')) { $line = $line.Substring(1) }
        $core = ConvertTo-GitRegex $line
        $rx = if ($anch) { "^$core$" } else { "^(?:.*/)?$core$" }
        try { $re = New-Object System.Text.RegularExpressions.Regex($rx, $opt) } catch { continue }
        $out += [pscustomobject]@{ Base = $baseDir; Neg = $neg; DirOnly = $dirOnly; Rx = $re; Src = $raw.Trim(); File = $label }
    }
    return $out
}

function Test-GitRules($rules, [string]$rel, [bool]$isDir) {
    $ign = $false
    foreach ($r in $rules) {
        $sub = $rel
        if ($r.Base) {
            if (-not $rel.StartsWith($r.Base)) { continue }
            $sub = $rel.Substring($r.Base.Length)
        }
        if ($r.DirOnly -and -not $isDir) { continue }
        if ($r.Rx.IsMatch($sub)) { $ign = -not $r.Neg }
    }
    return $ign
}

# true if the path itself or any parent folder is ignored (git never re-includes below an ignored folder)
function Test-GitIgnored($set, [string]$rel, [bool]$isDir = $false) {
    $parts = $rel.Split('/')
    $acc = ''
    for ($i = 0; $i -lt $parts.Length - 1; $i++) {
        $acc = if ($acc) { "$acc/$($parts[$i])" } else { $parts[$i] }
        if ($set.Cache.ContainsKey($acc)) { $v = $set.Cache[$acc] }
        else { $v = Test-GitRules $set.Rules $acc $true; $set.Cache[$acc] = $v }
        if ($v) { return $true }
    }
    return (Test-GitRules $set.Rules $rel $isDir)
}

function Get-GitIgnoreRules([string]$root, [string[]]$names, [bool]$withInfo = $true) {
    $icase = ($env:OS -eq 'Windows_NT')
    $set = [pscustomobject]@{ Rules = @(); Files = @(); Cache = @{} }
    $info = Join-Path $root '.git/info/exclude'
    if ($withInfo -and (Test-Path $info -PathType Leaf)) {
        $set.Rules += @(Read-GitIgnoreFile $info '' '.git/info/exclude' $icase); $set.Files += '.git/info/exclude'
    }
    $q = New-Object System.Collections.Queue
    $q.Enqueue($root)
    while ($q.Count -gt 0) {
        $d = $q.Dequeue()
        $dirRel = Get-Rel $root $d
        $baseDir = if ($dirRel) { "$dirRel/" } else { '' }
        foreach ($nm in $names) {
            $f = Join-Path $d $nm
            if (Test-Path $f -PathType Leaf) {
                $label = if ($dirRel) { "$dirRel/$nm" } else { $nm }
                $got = @(Read-GitIgnoreFile $f $baseDir $label $icase)
                if ($got.Count -gt 0) { $set.Rules += $got }
                $set.Files += $label
                $set.Cache = @{}
            }
        }
        foreach ($sd in [System.IO.Directory]::GetDirectories($d)) {
            $leaf = [System.IO.Path]::GetFileName($sd)
            if ($script:SkipDirs -contains $leaf) { continue }
            $srel = if ($dirRel) { "$dirRel/$leaf" } else { $leaf }
            if ($set.Rules.Count -gt 0 -and (Test-GitIgnored $set $srel $true)) { continue }   # prune ignored folders
            $q.Enqueue($sd)
        }
    }
    $set.Cache = @{}
    return $set
}

function Get-GitFileSet([string]$root) {
    if (-not (Get-Command git -ErrorAction SilentlyContinue)) { return $null }
    $old = [Console]::OutputEncoding
    try {
        [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
        $lines = & git -C $root -c core.quotepath=off ls-files --cached --others --exclude-standard 2>$null
        if ($LASTEXITCODE -ne 0) { return $null }
    } catch { return $null } finally { try { [Console]::OutputEncoding = $old } catch {} }
    $hs = @{}
    foreach ($l in $lines) { if ($l) { $hs[([string]$l).Replace('\','/')] = $true } }
    return $hs
}


# Returns $null (no filtering) or an object with Kind = 'rules' | 'git' and a one-line Info.
function Get-IgnoreFilter([string]$root, [string]$gm = '') {
    Initialize-Excludes
    if (-not $gm) { $gm = $GitIgnore }
    if ($gm -eq 'off') { return $null }
    $names = @($IgnoreFiles | Where-Object { $_ })
    if ($gm -eq 'git') {
        $gs = Get-GitFileSet $root
        if ($gs) {
            $extra = $null
            $others = @($names | Where-Object { $_ -ne '.gitignore' })
            if ($others.Count -gt 0) { $x = Get-GitIgnoreRules $root $others $false; if ($x.Rules.Count -gt 0) { $extra = $x } }
            $more = if ($extra) { " + $($extra.Rules.Count) rule(s) from $($others -join ', ')" } else { '' }
            return [pscustomobject]@{ Kind = 'git'; GitSet = $gs; Set = $extra; Info = "ignore: real git ($($gs.Count) tracked / untracked-not-ignored files)$more" }
        }
        Write-Warning "git mode unavailable (git missing or not a repository) - using the built-in parser"
    }
    if ($names.Count -eq 0) { return $null }
    $set = Get-GitIgnoreRules $root $names $true
    if ($set.Rules.Count -eq 0) { return $null }
    return [pscustomobject]@{ Kind = 'rules'; GitSet = $null; Set = $set; Info = "ignore files applied: $($set.Rules.Count) rule(s) from $($set.Files.Count) file(s)" }
}

function Test-Ignored($flt, [string]$rel) {
    if (-not $flt) { return $false }
    if ($flt.Kind -eq 'git') {
        if (-not $flt.GitSet.ContainsKey($rel)) { return $true }
        if ($flt.Set) { return (Test-GitIgnored $flt.Set $rel $false) }
        return $false
    }
    return (Test-GitIgnored $flt.Set $rel $false)
}
function Test-DirIgnored($flt, [string]$rel) {
    if (-not $flt -or -not $flt.Set) { return $false }
    return (Test-GitIgnored $flt.Set $rel $true)
}

# ============================================================
# scan: the one place that decides which files and folders are part of a bundle
# ============================================================
# does a path match -ForceInclude?  A plain name (node_modules, .env) matches at any depth; a pattern with / or * matches the path.
function Test-Forced([string]$rel) {
    foreach ($x in $ForceInclude) {
        if ($x -notmatch '[*?/]') { if (($rel -split '/') -contains $x) { return $true } }
        elseif ($rel -like $x -or $rel -like "$x/*" -or ("$rel/") -like $x) { return $true }
    }
    return $false
}

function Test-FileSkipped([string]$name, [string]$rel, [string]$full, [string]$self, [datetime]$mtimeUtc, [bool]$forced = $false) {
    if ($self -and ($full -ieq $self)) { return $true }
    foreach ($pat in $alwaysExcludeFiles) { if ($name -like $pat) { return $true } }
    if (-not $NoDefaultExcludes -and -not $forced) { foreach ($pat in $defaultExcludeFiles) { if ($name -like $pat) { return $true } } }
    if ($SkipIgnoreFiles) { if ((@($IgnoreFiles) + @('.gitattributes','.gitmodules','.gitkeep')) -contains $name) { return $true } }
    foreach ($x in $Exclude) { if ($rel -like $x -or $name -like $x -or $rel -like "$x/*") { return $true } }
    if ($Include.Count -gt 0) {
        $ok = $false
        foreach ($x in $Include) { if ($rel -like $x -or $name -like $x -or $rel -like "$x/*") { $ok = $true; break } }
        if (-not $ok) { return $true }
    }
    if ($OnlyExt.Count -gt 0) {
        $ext = [System.IO.Path]::GetExtension($name).TrimStart('.').ToLowerInvariant()
        if ($OnlyExt -notcontains $ext) { return $true }
    }
    if ($Since -ne [datetime]::MinValue -and $mtimeUtc -lt $Since.ToUniversalTime()) { return $true }
    return $false
}

# .git is never packed; -ExcludeDir always wins; -ForceInclude beats the default excludes
function Test-DirSkipped([string]$leaf, [string]$rel, [bool]$forced = $false) {
    if ($alwaysExcludeDirs -contains $leaf) { return $true }
    foreach ($x in $ExcludeDir) { if ($leaf -like $x -or $rel -like $x -or $rel -like "$x/*") { return $true } }
    if ($forced) { return $false }
    return ($script:SkipDirs -contains $leaf)
}

# Files = would be packed.  IgnoredFiles / IgnoredDirs = removed by the ignore files.
# SkippedDirs = left out by the default excludes / -ExcludeDir.  EmptyDirs = empty folders.
function Get-SourceScan([string]$root, $flt, [bool]$measure = $true, [string]$self = '') {
    Initialize-Excludes
    $files  = New-Object System.Collections.ArrayList
    $ignF   = New-Object System.Collections.ArrayList
    $ignD   = New-Object System.Collections.ArrayList
    $skipD  = New-Object System.Collections.ArrayList
    $emptyD = New-Object System.Collections.ArrayList
    $q = New-Object System.Collections.Queue
    $q.Enqueue($root)
    while ($q.Count -gt 0) {
        $d = $q.Dequeue()
        $dirRel = Get-Rel $root $d
        $subs = [System.IO.Directory]::GetDirectories($d)
        $fps  = [System.IO.Directory]::GetFiles($d)
        if (($subs.Count + $fps.Count) -eq 0 -and $dirRel -and (-not $flt -or $flt.Kind -ne 'git')) { [void]$emptyD.Add($dirRel) }
        foreach ($sd in $subs) {
            $leaf = [System.IO.Path]::GetFileName($sd)
            $srel = if ($dirRel) { "$dirRel/$leaf" } else { $leaf }
            if (([System.IO.File]::GetAttributes($sd) -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { continue }   # never follow links
            if ($alwaysExcludeDirs -contains $leaf) { continue }
            $forced = Test-Forced $srel
            $isSkipped = Test-DirSkipped $leaf $srel $forced
            $isIgnored = (-not $forced) -and (-not $isSkipped) -and (Test-DirIgnored $flt $srel)
            if ($isSkipped -or $isIgnored) {
                $cnt = 0; $bytes = [long]0
                if ($measure) {
                    $m = Get-ChildItem -Path $sd -Recurse -File -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum
                    $cnt = [int]$m.Count; if ($m.Sum) { $bytes = [long]$m.Sum }
                }
                $row = [pscustomobject]@{ Rel = $srel; Count = $cnt; Size = $bytes }
                if ($isSkipped) { [void]$skipD.Add($row) } else { [void]$ignD.Add($row) }
                continue
            }
            $q.Enqueue($sd)
        }
        foreach ($fp in $fps) {
            $fi = New-Object System.IO.FileInfo $fp
            $rel = if ($dirRel) { "$dirRel/$($fi.Name)" } else { $fi.Name }
            if ($fi.LinkTarget) { continue }   # never follow file symlinks (they may point outside the project)
            $forced = Test-Forced $rel
            if (Test-FileSkipped $fi.Name $rel $fi.FullName $self $fi.LastWriteTimeUtc $forced) { continue }
            $item = [pscustomobject]@{ Rel = $rel; Size = [long]$fi.Length; Full = $fi.FullName; MTime = $fi.LastWriteTime }
            if (-not $forced -and (Test-Ignored $flt $rel)) { [void]$ignF.Add($item) } else { [void]$files.Add($item) }
        }
    }
    return [pscustomobject]@{ Files = $files; IgnoredFiles = $ignF; IgnoredDirs = $ignD; SkippedDirs = $skipD; EmptyDirs = $emptyD }
}

# ============================================================
# pack / watch
# ============================================================
function Apply-Profile {
    switch ($Profile) {
        'quick' { $script:Exclude = @($Exclude) + @('*.log','*.bak'); $script:NoBinary = [switch]$true }
        'code'  { $script:OnlyExt = @('go','py','js','ts','tsx','jsx','json','md','yaml','yml','ps1','sh','html','css','sql','proto','toml','ini','conf','xml','txt','cs','java','rs','c','h','cpp'); $script:NoBinary = [switch]$true }
        'docs'  { $script:OnlyExt = @('md','txt','rst','json','yaml','yml'); $script:NoBinary = [switch]$true }
    }
}

# shared by pack -ExcludeSecrets and the secrets command
$script:SecretNameRx = '(^|/)(\.env(\..*)?|id_rsa|id_dsa|id_ed25519|.*\.pem|.*\.pfx|.*\.p12|.*\.keystore|credentials\.json|\.npmrc|\.pypirc|\.netrc)$'
$script:SecretRules = @(
    @{ N = 'private key block';  R = '-----BEGIN (RSA |EC |DSA |OPENSSH |PGP )?PRIVATE KEY' }
    @{ N = 'AWS access key id';  R = '\bAKIA[0-9A-Z]{16}\b' }
    @{ N = 'GitHub token';       R = '\bgh[pousr]_[A-Za-z0-9]{36,}\b' }
    @{ N = 'Slack token';        R = '\bxox[baprs]-[A-Za-z0-9-]{10,}' }
    @{ N = 'Google API key';     R = '\bAIza[0-9A-Za-z_\-]{35}\b' }
    @{ N = 'JWT';                R = '\beyJ[A-Za-z0-9_-]{10,}\.eyJ[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}' }
    @{ N = 'credentials in URL'; R = '[a-z][a-z0-9+.-]*://[^/\s:@''"]+:[^/\s@''"]{3,}@' }
    @{ N = 'hard-coded secret';  R = '(?i)(password|passwd|secret|api[_-]?key|access[_-]?token|auth[_-]?token)\s*[:=]\s*[''"][^''"\s]{8,}[''"]' }
)
function Test-SecretFile([string]$full, [string]$rel, [long]$size) {
    if ($rel -match $script:SecretNameRx -and $rel -notmatch '\.(example|sample|template)$') { return 'sensitive file name' }
    if ($size -le 0 -or $size -gt 1MB) { return $null }
    $bytes = [System.IO.File]::ReadAllBytes($full)
    if (Test-BytesBinary $bytes) { return $null }
    foreach ($line in $utf8.GetString($bytes).Split("`n")) {
        if ($line.Length -gt 2000) { continue }
        foreach ($r in $script:SecretRules) { if ([regex]::IsMatch($line, $r.R)) { return $r.N } }
    }
    return $null
}

# files git reports as changed (working tree + untracked, or changed since a ref), relative to $root
function Get-ChangedSet([string]$root) {
    if (-not (Get-Command git -ErrorAction SilentlyContinue)) { throw "git is required for -Changed / -ChangedSince" }
    $old = [Console]::OutputEncoding
    $set = @{}
    try {
        [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
        $prefix = (& git -C $root rev-parse --show-prefix 2>$null)
        if ($LASTEXITCODE -ne 0) { throw "$root is not inside a git repository" }
        $prefix = "$prefix".Trim()
        foreach ($l in (& git -C $root -c core.quotepath=off status --porcelain -uall 2>$null)) {
            if ($l.Length -le 3) { continue }
            $p = $l.Substring(3); if ($p -match ' -> ') { $p = ($p -split ' -> ')[-1] }
            $p = $p.Trim('"').Replace('\','/')
            if ($prefix -and $p.StartsWith($prefix)) { $p = $p.Substring($prefix.Length) } elseif ($prefix) { continue }
            $set[$p] = $true
        }
        if ($ChangedSince) {
            foreach ($p in (& git -C $root -c core.quotepath=off diff --name-only --relative $ChangedSince 2>$null)) { if ($p) { $set[([string]$p).Replace('\','/')] = $true } }
        }
    } finally { try { [Console]::OutputEncoding = $old } catch {} }
    return $set
}

# write ONE bundle file for a group of files; returns counters
function Write-PackFile($files, $dirs, [string]$dest, [string]$name, [string]$baseName) {
    $tmp = "$dest.$([guid]::NewGuid().ToString('N').Substring(0,8)).sbtmp"
    $w = New-Object System.IO.StreamWriter($tmp, $false, $utf8); $w.NewLine = "`n"
    $files = @($files | Where-Object { $_ }); $dirs = @($dirs | Where-Object { $_ })
    $total = @($files).Count; $i = 0; $skipped = 0; $bytesIn = [long]0; $textBytes = [long]0
    try {
        Write-BundleHeader $w $name $total @($dirs).Count $Compress.IsPresent $files $dirs $baseName
        foreach ($f in $files) {
            $i++
            $e = New-FileEntry $f.Full $f.Rel
            if ($e.Enc -eq 'skipped') { $skipped++ } else { $bytesIn += $e.Size; if ($e.Enc -eq 'utf8') { $textBytes += $e.Size } }
            Write-BundleEntry $w $e
            if ($i % 25 -eq 0 -or $i -eq $total) {
                Write-Host ("`r  [{0}/{1}] {2,5}%  {3}" -f $i,$total,[math]::Round(100 * $i / [Math]::Max(1,$total),0),(Format-Size $bytesIn)).PadRight(60) -NoNewline
            }
        }
        Write-BundleFooter $w $dirs
    } catch {
        try { $w.Dispose() } catch {}
        if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
        throw
    }
    $w.Dispose()
    Complete-BundleFile $tmp $dest $Compress.IsPresent
    Write-Host ""
    return @{ Files = $total; Dirs = @($dirs).Count; Bytes = $bytesIn; TextBytes = $textBytes; Skipped = $skipped; Dest = $dest; OnDisk = (Get-Item $dest).Length }
}

function Invoke-Pack {
    Apply-Profile
    $root = (Resolve-Path $Source).Path
    $dest = Get-AbsPath $BundleFile
    $destDir = [System.IO.Path]::GetDirectoryName($dest)
    if (-not (Test-Path $destDir)) { New-Item -ItemType Directory -Path $destDir -Force | Out-Null }
    Write-Host "Packing: $root" -ForegroundColor Cyan
    $flt = Get-IgnoreFilter $root
    if ($flt) { Write-Host "  $($flt.Info)" -ForegroundColor Cyan } else { Write-Host "  ignore files: not applied (everything is included)" -ForegroundColor DarkGray }
    $scan = Get-SourceScan $root $flt $false $dest
    $files = @($scan.Files | Sort-Object Rel)
    $dirs = if ($NoEmptyDirs) { @() } else { @($scan.EmptyDirs | Sort-Object) }
    $ignoredN = $scan.IgnoredFiles.Count
    $baseName = ''

    if ($ExcludeSecrets) {
        $kept = @(); $excl = @()
        foreach ($f in $files) { $why = Test-SecretFile $f.Full $f.Rel $f.Size; if ($why) { $excl += "$($f.Rel)  [$why]" } else { $kept += $f } }
        $files = $kept
        if ($excl.Count -gt 0) { Write-Host "  left out ($($excl.Count)) because they look sensitive:" -ForegroundColor Yellow; $excl | Select-Object -First 15 | ForEach-Object { Write-Host "    $_" -ForegroundColor Yellow } }
    }
    if ($Changed -or $ChangedSince) {
        $cs = Get-ChangedSet $root
        $files = @($files | Where-Object { $cs.ContainsKey($_.Rel) }); $dirs = @()
        Write-Host "  git changes only: $($files.Count) file(s)" -ForegroundColor Cyan
    }
    if ($Base) {
        $bd = Read-Bundle (Resolve-Path $Base).Path $Key
        $bh = @{}; foreach ($e in $bd.Entries) { $bh[$e.Path] = $e.Hash }
        $files = @($files | Where-Object { -not $bh.ContainsKey($_.Rel) -or (Get-Sha256File $_.Full) -ne $bh[$_.Rel] }); $dirs = @()
        $baseName = [System.IO.Path]::GetFileName($Base)
        Write-Host "  incremental vs $baseName : $($files.Count) new/changed file(s)" -ForegroundColor Cyan
    }

    if ($ForceInclude.Count -gt 0) {
        $forcedN = @($files | Where-Object { Test-Forced $_.Rel }).Count
        Write-Host ("  force-included: {0}  ({1} file(s))" -f ($ForceInclude -join ', '),$forcedN) -ForegroundColor Cyan
    }

    # preview only: show what WOULD be packed
    if ($DryRun) {
        $shown = 0; $sumB = [long]0
        foreach ($f in $files) { $sumB += $f.Size; if ($shown -lt 400) { Write-Host ("  {0,10}  {1}" -f (Format-Size $f.Size),$f.Rel); $shown++ } }
        if ($files.Count -gt $shown) { Write-Host "  ... and $($files.Count - $shown) more" -ForegroundColor DarkGray }
        foreach ($d in $dirs) { Write-Host ("  {0,10}  {1}/" -f 'folder',$d) }
        Write-Host ""
        Write-Host ("PREVIEW: {0} file(s), {1} empty folder(s), {2}  (ignored: {3} file(s), {4} folder(s); skipped by default: {5} folder(s))" -f $files.Count,$dirs.Count,(Format-Size $sumB),$ignoredN,$scan.IgnoredDirs.Count,$scan.SkippedDirs.Count) -ForegroundColor Green
        Write-Host "  (preview only - nothing was written)" -ForegroundColor Yellow
        return
    }

    # split into parts (each part is a complete bundle on its own)
    $groups = @()
    if ($SplitMB -gt 0 -and $files.Count -gt 0) {
        $limit = [long]($SplitMB * 1MB); $cur = @(); $sum = [long]0
        foreach ($f in $files) {
            if ($cur.Count -gt 0 -and ($sum + $f.Size) -gt $limit) { $groups += ,$cur; $cur = @(); $sum = [long]0 }
            $cur += $f; $sum += $f.Size
        }
        if ($cur.Count -gt 0) { $groups += ,$cur }
    } else { $groups = @(,$files) }

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $name = [System.IO.Path]::GetFileName($root)
    $results = @()
    for ($g = 0; $g -lt $groups.Count; $g++) {
        $partDest = $dest
        if ($groups.Count -gt 1) {
            $partDest = Join-Path $destDir ("{0}.part{1}{2}" -f [System.IO.Path]::GetFileNameWithoutExtension($dest),($g + 1),[System.IO.Path]::GetExtension($dest))
            Write-Host ("  part {0}/{1}" -f ($g + 1),$groups.Count) -ForegroundColor Cyan
        }
        $pd = @(); if ($g -eq $groups.Count - 1) { $pd = @($dirs) }
        $results += Write-PackFile $groups[$g] $pd $partDest $name $baseName
    }
    $sw.Stop()
    $tf = 0; $td = 0; $tb = [long]0; $tt = [long]0; $ts = 0; $disk = [long]0
    foreach ($r in $results) { $tf += $r.Files; $td += $r.Dirs; $tb += $r.Bytes; $tt += $r.TextBytes; $ts += $r.Skipped; $disk += $r.OnDisk }
    Write-Host ("OK: {0} file(s), {1} empty folder(s), {2} content -> {3}  ({4:N1}s)" -f $tf,$td,(Format-Size $tb),(Format-Size $disk),$sw.Elapsed.TotalSeconds) -ForegroundColor Green
    Write-Host ("  ~{0:N0} tokens of text (rough estimate: 1 token ~ 4 characters)" -f ($tt / 4)) -ForegroundColor DarkGray
    if ($flt) { Write-Host ("  ignored by ignore files: {0} file(s), {1} folder(s)" -f $ignoredN,$scan.IgnoredDirs.Count) -ForegroundColor DarkGray }
    if ($ts -gt 0) { Write-Host "  $ts file(s) listed but not stored (binary skipped / too large)" -ForegroundColor Yellow }
    foreach ($r in $results) { Write-Host "  => $($r.Dest)" }
    if ($Clipboard) {
        if ($results.Count -ne 1 -or $Compress) { Write-Warning "-Clipboard needs a single, uncompressed bundle (no -SplitMB / -Compress)" }
        elseif (-not (Get-Command Set-Clipboard -ErrorAction SilentlyContinue)) { Write-Warning "Set-Clipboard is not available on this system" }
        else {
            try { Set-Clipboard -Value ([System.IO.File]::ReadAllText($results[0].Dest, $utf8)); Write-Host "  copied to the clipboard" -ForegroundColor Green }
            catch { Write-Warning "could not copy to the clipboard: $($_.Exception.Message)" }
        }
    }
}

function Get-ScanSignature([string]$root, $flt, [string]$dest) {
    $scan = Get-SourceScan $root $flt $false $dest
    $sb = New-Object System.Text.StringBuilder
    foreach ($f in ($scan.Files | Sort-Object Rel)) { [void]$sb.Append($f.Rel).Append('|').Append($f.Size).Append('|').Append($f.MTime.Ticks).Append("`n") }
    return (Get-Sha256 ($utf8.GetBytes($sb.ToString())))
}

function Invoke-Watch {
    $root = (Resolve-Path $Source).Path
    $dest = Get-AbsPath $BundleFile
    Write-Host "Watching $root every ${IntervalSec}s - Ctrl+C to stop" -ForegroundColor Cyan
    Invoke-Pack
    $sig = Get-ScanSignature $root (Get-IgnoreFilter $root) $dest
    while ($true) {
        Start-Sleep -Seconds $IntervalSec
        $now = Get-ScanSignature $root (Get-IgnoreFilter $root) $dest
        if ($now -ne $sig) { Write-Host "`nchange detected - packing again" -ForegroundColor Yellow; Invoke-Pack; $sig = $now }
    }
}

# ============================================================
# restore: unpack / sync / extract
# ============================================================
function Restore-Entries($data, [string]$root, [bool]$dry, [bool]$ow, [string]$flt, [bool]$deleteExtras) {
    $done = 0; $skip = 0; $conflict = 0; $unchanged = 0; $unsafe = 0; $bytes = [long]0
    $written = @{}
    $sel = @($data.Entries | Where-Object { $_.Path -like $flt })
    $i = 0
    foreach ($e in $sel) {
        $i++
        $written[$e.Path] = $true
        if ($e.Enc -eq 'skipped') { $skip++; continue }
        $out = Join-Safe $root $e.Path
        if (-not $out) { Write-Warning "refusing unsafe path in bundle: $($e.Path)"; $unsafe++; continue }
        $exists = Test-Path -LiteralPath $out -PathType Leaf
        if ($exists -and $e.Hash -and ((Get-Sha256File $out) -eq $e.Hash)) { $unchanged++; continue }
        if ($exists -and -not $ow) {
            if ($dry) { Write-Host "  [dry] CONFLICT (exists, differs): $($e.Path)" -ForegroundColor Yellow }
            else { Write-Warning "skip (exists, differs - use -Overwrite): $($e.Path)" }
            $conflict++; continue
        }
        if ($dry) { Write-Host "  [dry] $($e.Path)  ($(Format-Size ([long]$e.Size)))"; $done++; continue }
        $dir = [System.IO.Path]::GetDirectoryName($out)
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        [System.IO.File]::WriteAllBytes($out, (Get-EntryBytes $e))
        if ($PreserveTimes -and $e.MTime -gt 0) {
            try { [System.IO.File]::SetLastWriteTimeUtc($out, [DateTimeOffset]::FromUnixTimeSeconds([long]$e.MTime).UtcDateTime) } catch {}
        }
        $done++; $bytes += [long]$e.Size
        if ($i % 25 -eq 0 -or $i -eq $sel.Count) { Write-Host ("`r  [{0}/{1}] {2}" -f $i,$sel.Count,$e.Path).PadRight(70) -NoNewline }
    }
    $dirsMade = 0
    foreach ($d in $data.Dirs) {
        if ($d -notlike $flt -and $flt -ne '*') { continue }
        $out = Join-Safe $root $d
        if (-not $out) { $unsafe++; continue }
        $written[$d] = $true
        if (-not (Test-Path -LiteralPath $out)) {
            if ($dry) { Write-Host "  [dry] empty folder: $d" } else { New-Item -ItemType Directory -Path $out -Force | Out-Null }
            $dirsMade++
        }
    }
    if ($deleteExtras -and -not $dry) {
        Initialize-Excludes
        $igf = if ($GitIgnore -ne 'off') { Get-IgnoreFilter $root 'auto' } else { $null }   # never delete files the ignore files protect
        Get-ChildItem -Path $root -Recurse -File -Force | ForEach-Object {
            $rel = Get-Rel $root $_.FullName
            $protect = $false
            foreach ($d in $script:SkipDirs) { if (($rel -split '/') -contains $d) { $protect = $true; break } }
            if (-not $protect -and $igf -and (Test-Ignored $igf $rel)) { $protect = $true }
            if (-not $protect -and -not $written.ContainsKey($rel)) { Write-Warning "delete: $rel"; Remove-Item -LiteralPath $_.FullName -Force }
        }
    }
    Write-Host ""
    Write-Host ("OK: {0} restored, {1} unchanged, {2} empty folder(s), {3} skipped, {4} conflicts, {5}" -f $done,$unchanged,$dirsMade,$skip,$conflict,(Format-Size $bytes)) -ForegroundColor Green
    if ($unsafe -gt 0) { Write-Host "  $unsafe unsafe path(s) were refused" -ForegroundColor Red }
}

function Invoke-Unpack {
    $bp = (Get-BundleInput)
    $data = Read-Bundle $bp $Key
    if (-not $data.Complete) { Write-Warning "bundle has no end marker - it may be truncated" }
    if (-not (Test-Path $Target) -and -not $DryRun) { New-Item -ItemType Directory -Path $Target -Force | Out-Null }
    $tr = if (Test-Path $Target) { (Resolve-Path $Target).Path } else { Get-AbsPath $Target }
    Restore-Entries $data $tr $DryRun.IsPresent $Overwrite.IsPresent $Filter $false
}
function Invoke-Sync {
    $bp = (Get-BundleInput)
    $data = Read-Bundle $bp $Key
    if (-not (Test-Path $Target)) { New-Item -ItemType Directory -Path $Target -Force | Out-Null }
    Restore-Entries $data (Resolve-Path $Target).Path $DryRun.IsPresent $true $Filter $DeleteExtras.IsPresent
}
function Invoke-Extract {
    if (-not $Path) { throw "-Path is required for extract" }
    $data = Read-Bundle (Get-BundleInput) $Key
    $prefix = ($Path -replace '\\','/').Trim('/')
    $sel = @($data.Entries | Where-Object { $_.Path -eq $prefix -or $_.Path -like "$prefix/*" })
    $selDirs = @($data.Dirs | Where-Object { $_ -eq $prefix -or $_ -like "$prefix/*" })
    if ($sel.Count -eq 0 -and $selDirs.Count -eq 0) { Write-Warning "not found in bundle: $Path"; return }
    if (-not (Test-Path $Target) -and -not $DryRun) { New-Item -ItemType Directory -Path $Target -Force | Out-Null }
    $tr = if (Test-Path $Target) { (Resolve-Path $Target).Path } else { Get-AbsPath $Target }
    Restore-Entries @{ Entries = $sel; Dirs = $selDirs } $tr $DryRun.IsPresent $Overwrite.IsPresent '*' $false
}

# ============================================================
# inspect: info / list / tree / stats / grep / verify / duplicates
# ============================================================
function Write-PathTree($items, $dirs, [int]$maxDepth, [bool]$withSize = $true) {
    $rows = @()
    foreach ($it in $items) { $rows += [pscustomobject]@{ Rel = $it.Rel; Size = $it.Size; Dir = $false } }
    foreach ($d in $dirs)   { $rows += [pscustomobject]@{ Rel = $d; Size = 0; Dir = $true } }
    $prev = @()
    foreach ($r in ($rows | Sort-Object Rel)) {
        $parts = $r.Rel.Split('/')
        $depthDirs = if ($r.Dir) { $parts.Count } else { $parts.Count - 1 }
        $common = 0
        while ($common -lt $prev.Count -and $common -lt $depthDirs -and $prev[$common] -eq $parts[$common]) { $common++ }
        for ($i = $common; $i -lt $depthDirs; $i++) {
            if ($maxDepth -eq 0 -or ($i + 1) -le $maxDepth) {
                $suffix = if ($r.Dir -and $i -eq $depthDirs - 1) { '/   (empty)' } else { '/' }
                Write-Host (('  ' * $i) + $parts[$i] + $suffix)
            }
        }
        if (-not $r.Dir -and ($maxDepth -eq 0 -or $parts.Count -le $maxDepth)) {
            $sz = if ($withSize) { '  (' + (Format-Size $r.Size) + ')' } else { '' }
            Write-Host (('  ' * ($parts.Count - 1)) + $parts[$parts.Count - 1] + $sz)
        }
        if ($depthDirs -gt 0) { $prev = @($parts[0..($depthDirs - 1)]) } else { $prev = @() }
    }
}

function Invoke-Info {
    $bp = (Get-BundleInput)
    $data = Read-Bundle $bp $Key
    $h = $data.Header; $en = $data.Entries
    $by = @{}; $total = [long]0
    foreach ($e in $en) { $by[$e.Enc] = 1 + [int]$by[$e.Enc]; $total += [long]$e.Size }
    if ($OutputJson) {
        @{ format = "$script:FormatName v$($h.Version)"; created = $h.Created; name = $h.Name; files = $en.Count; empty_dirs = $data.Dirs.Count; content_bytes = $total; on_disk = (Get-Item $bp).Length; complete = $data.Complete } | ConvertTo-Json
        return
    }
    Write-Host "Bundle:   $bp"
    Write-Host "Format:   $script:FormatName v$($h.Version)    Created: $($h.Created)    Name: $($h.Name)"
    $enc = ($by.Keys | Sort-Object | ForEach-Object { "$_ $($by[$_])" }) -join ', '
    Write-Host ("Files:    {0}  ({1})" -f $en.Count,$enc)
    Write-Host ("Folders:  {0} empty folder(s) recorded" -f $data.Dirs.Count)
    Write-Host ("Content:  {0} original,  {1} on disk" -f (Format-Size $total),(Format-Size (Get-Item $bp).Length))
    $tx = [long]0; foreach ($e in $en) { if ($e.Enc -eq 'utf8') { $tx += [long]$e.Size } }
    Write-Host ("Tokens:   ~{0:N0} (text only, rough estimate)" -f ($tx / 4))
    if ($h.Base) { Write-Host "Base:     incremental bundle, relative to $($h.Base)" }
    Write-Host ("Complete: {0}" -f $(if ($data.Complete) { 'yes' } else { 'NO - end marker missing (truncated?)' }))
}

function Invoke-List {
    $data = Read-Bundle (Get-BundleInput) $Key
    $rows = $data.Entries | Sort-Object Path
    if ($OutputJson) { $rows | Select-Object Path,Enc,Size,MTime,Hash,Eol | ConvertTo-Json; return }
    foreach ($e in $rows) {
        $tag = if ($e.Enc -eq 'utf8') { "utf8/$($e.Eol)" } elseif ($e.Enc -eq 'skipped') { "skipped" } else { $e.Enc }
        if ($Long) {
            $mt = if ($e.MTime -gt 0) { [DateTimeOffset]::FromUnixTimeSeconds([long]$e.MTime).LocalDateTime.ToString('yyyy-MM-dd HH:mm') } else { '-' }
            $hs = if ($e.Hash) { $e.Hash.Substring(0,10) } else { '-' }
            Write-Host ("{0,10}  {1,-10} {2}  {3}  {4}" -f (Format-Size ([long]$e.Size)),$tag,$mt,$hs,$e.Path)
        } else { Write-Host ("{0,10}  {1,-10} {2}" -f (Format-Size ([long]$e.Size)),$tag,$e.Path) }
    }
    foreach ($d in ($data.Dirs | Sort-Object)) { Write-Host ("{0,10}  {1,-10} {2}/" -f '-','folder',$d) }
}

function Invoke-Tree {
    $data = Read-Bundle (Get-BundleInput) $Key
    $items = @($data.Entries | ForEach-Object { [pscustomobject]@{ Rel = $_.Path; Size = [long]$_.Size } })
    Write-PathTree $items @($data.Dirs) $Depth $true
    Write-Host ""
    Write-Host ("{0} file(s), {1} empty folder(s)" -f $items.Count,$data.Dirs.Count) -ForegroundColor Green
}

function Invoke-Stats {
    $data = Read-Bundle (Get-BundleInput) $Key
    $byExt = @{}; $byTop = @{}
    foreach ($e in $data.Entries) {
        $ext = [System.IO.Path]::GetExtension($e.Path).ToLowerInvariant(); if (-not $ext) { $ext = '(none)' }
        $topDir = if ($e.Path.Contains('/')) { $e.Path.Substring(0, $e.Path.IndexOf('/')) } else { '(root)' }
        if (-not $byExt.ContainsKey($ext)) { $byExt[$ext] = @{ N = 0; B = [long]0 } }
        if (-not $byTop.ContainsKey($topDir)) { $byTop[$topDir] = @{ N = 0; B = [long]0 } }
        $byExt[$ext].N++; $byExt[$ext].B += [long]$e.Size
        $byTop[$topDir].N++; $byTop[$topDir].B += [long]$e.Size
    }
    if ($OutputJson) {
        @{ by_extension = ($byExt.Keys | ForEach-Object { @{ ext = $_; files = $byExt[$_].N; bytes = $byExt[$_].B } }); by_folder = ($byTop.Keys | ForEach-Object { @{ folder = $_; files = $byTop[$_].N; bytes = $byTop[$_].B } }) } | ConvertTo-Json -Depth 4
        return
    }
    Write-Host "By extension:" -ForegroundColor Yellow
    $byExt.Keys | Sort-Object { $byExt[$_].B } -Descending | Select-Object -First 15 | ForEach-Object { Write-Host ("  {0,-10} {1,6} files  {2}" -f $_,$byExt[$_].N,(Format-Size $byExt[$_].B)) }
    Write-Host "By top-level folder:" -ForegroundColor Yellow
    $byTop.Keys | Sort-Object { $byTop[$_].B } -Descending | Select-Object -First 15 | ForEach-Object { Write-Host ("  {0,-24} {1,6} files  {2}" -f $_,$byTop[$_].N,(Format-Size $byTop[$_].B)) }
}

function Invoke-Grep {
    if (-not $Pattern) { throw "-Pattern is required" }
    $data = Read-Bundle (Get-BundleInput) $Key
    $opt = if ($CaseSensitive) { [System.Text.RegularExpressions.RegexOptions]::None } else { [System.Text.RegularExpressions.RegexOptions]::IgnoreCase }
    $rx = New-Object System.Text.RegularExpressions.Regex($Pattern, $opt)
    $hits = 0; $fileHits = 0
    foreach ($e in ($data.Entries | Sort-Object Path)) {
        if ($e.Enc -ne 'utf8') { continue }
        $lines = $e.Content.Split("`n")
        $n = $lines.Count; if ($n -gt 0 -and $lines[$n - 1] -eq '') { $n-- }
        $shown = @{}; $any = $false
        for ($i = 0; $i -lt $n; $i++) {
            if (-not $rx.IsMatch($lines[$i])) { continue }
            $hits++; $any = $true
            $from = [Math]::Max(0, $i - $Context); $to = [Math]::Min($n - 1, $i + $Context)
            for ($j = $from; $j -le $to; $j++) {
                if ($shown.ContainsKey($j)) { continue }
                $shown[$j] = $true
                $mark = if ($j -eq $i) { ':' } else { '-' }
                Write-Host ("{0}{1}{2} {3}" -f $e.Path,$mark,($j + 1),$lines[$j].TrimEnd())
            }
            if ($Context -gt 0) { Write-Host "--" -ForegroundColor DarkGray }
        }
        if ($any) { $fileHits++ }
    }
    Write-Host ""
    Write-Host ("{0} match(es) in {1} file(s)" -f $hits,$fileHits) -ForegroundColor Green
}

function Invoke-Verify {
    $data = Read-Bundle (Get-BundleInput) $Key
    $ok = 0; $bad = 0; $noHash = 0; $skipped = 0; $unsafe = 0
    foreach ($e in $data.Entries) {
        if (-not (Test-SafeRel $e.Path)) { Write-Warning "unsafe path: $($e.Path)"; $unsafe++ }
        if ($e.Enc -eq 'skipped') { $skipped++; continue }
        if (-not $e.Hash) { $noHash++; continue }
        $actual = ''
        try { $actual = Get-Sha256 (Get-EntryBytes $e) } catch { $actual = 'unreadable' }
        if ($actual -eq $e.Hash) { $ok++ } else { $bad++; Write-Warning "MISMATCH: $($e.Path)" }
    }
    if (-not $data.Complete) { Write-Warning "end marker missing - bundle may be truncated"; $bad++ }
    if ([int]$data.Header.Files -ne $data.Entries.Count) { Write-Warning "header says $($data.Header.Files) files, found $($data.Entries.Count)"; $bad++ }
    Write-Host ""
    $col = if ($bad -gt 0 -or $unsafe -gt 0) { 'Red' } else { 'Green' }
    Write-Host ("OK={0}  BAD={1}  skipped={2}  no-hash={3}  unsafe-paths={4}" -f $ok,$bad,$skipped,$noHash,$unsafe) -ForegroundColor $col
    if ($bad -gt 0 -or $unsafe -gt 0) { $script:VerifyFailed = $true }
}

function Invoke-Duplicates {
    $data = Read-Bundle (Get-BundleInput) $Key
    $groups = $data.Entries | Where-Object { $_.Hash -and [long]$_.Size -gt 0 } | Group-Object Hash | Where-Object { $_.Count -gt 1 }
    if (-not $groups) { Write-Host "no duplicate files" -ForegroundColor Green; return }
    $waste = [long]0
    foreach ($g in $groups) {
        $sz = [long]$g.Group[0].Size; $waste += $sz * ($g.Count - 1)
        Write-Host ("{0} copies, {1} each:" -f $g.Count,(Format-Size $sz)) -ForegroundColor Yellow
        $g.Group | Sort-Object Path | ForEach-Object { Write-Host "  $($_.Path)" }
    }
    Write-Host ""
    Write-Host ("{0} duplicate group(s), {1} wasted" -f @($groups).Count,(Format-Size $waste)) -ForegroundColor Green
}

# ============================================================
# compare: bundle vs folder, bundle vs bundle
# ============================================================
function Invoke-Diff {
    $root = (Resolve-Path $CompareTo).Path
    $data = Read-Bundle (Get-BundleInput) $Key
    $flt = Get-IgnoreFilter $root
    $scan = Get-SourceScan $root $flt $false ''
    $inB = @{}; foreach ($e in $data.Entries) { $inB[$e.Path] = $e }
    $inD = @{}; foreach ($f in $scan.Files) { $inD[$f.Rel] = $f }
    $onlyB = @(); $onlyD = @(); $diff = @(); $same = 0
    foreach ($k in $inB.Keys) {
        if (-not $inD.ContainsKey($k)) { $onlyB += $k; continue }
        $e = $inB[$k]
        if (-not $e.Hash) { $same++; continue }
        if ((Get-Sha256File $inD[$k].Full) -ne $e.Hash) { $diff += $k } else { $same++ }
    }
    foreach ($k in $inD.Keys) { if (-not $inB.ContainsKey($k)) { $onlyD += $k } }
    Write-Host "== only in bundle ==" -ForegroundColor Yellow; $onlyB | Sort-Object | ForEach-Object { Write-Host "  - $_" }
    Write-Host "== only in folder ==" -ForegroundColor Yellow; $onlyD | Sort-Object | ForEach-Object { Write-Host "  + $_" }
    Write-Host "== different ==" -ForegroundColor Yellow;      $diff  | Sort-Object | ForEach-Object { Write-Host "  ~ $_" }
    Write-Host ""
    Write-Host ("same={0}  only-bundle={1}  only-folder={2}  different={3}" -f $same,$onlyB.Count,$onlyD.Count,$diff.Count) -ForegroundColor Green
}

function Invoke-BundleDiff {
    if (-not $Compare2) { throw "-Compare2 (second bundle) is required" }
    $a = Read-Bundle (Get-BundleInput) $Key
    $b = Read-Bundle (Resolve-Path $Compare2).Path $Key
    $ma = @{}; foreach ($e in $a.Entries) { $ma[$e.Path] = $e }
    $mb = @{}; foreach ($e in $b.Entries) { $mb[$e.Path] = $e }
    $added = @(); $removed = @(); $changedList = @(); $same = 0
    foreach ($k in $mb.Keys) { if (-not $ma.ContainsKey($k)) { $added += $k } }
    foreach ($k in $ma.Keys) {
        if (-not $mb.ContainsKey($k)) { $removed += $k; continue }
        if ($ma[$k].Hash -and $mb[$k].Hash -and $ma[$k].Hash -ne $mb[$k].Hash) { $changedList += $k } else { $same++ }
    }
    Write-Host ("A: {0}" -f (Split-Path $BundleFile -Leaf)) -ForegroundColor Cyan
    Write-Host ("B: {0}" -f (Split-Path $Compare2 -Leaf)) -ForegroundColor Cyan
    Write-Host "== added in B ==" -ForegroundColor Yellow;   $added   | Sort-Object | ForEach-Object { Write-Host "  + $_" }
    Write-Host "== removed in B ==" -ForegroundColor Yellow; $removed | Sort-Object | ForEach-Object { Write-Host "  - $_" }
    Write-Host "== changed ==" -ForegroundColor Yellow;      $changedList | Sort-Object | ForEach-Object { Write-Host "  ~ $_" }
    Write-Host ""
    Write-Host ("same={0}  added={1}  removed={2}  changed={3}" -f $same,$added.Count,$removed.Count,$changedList.Count) -ForegroundColor Green
}

# ============================================================
# merge
# ============================================================
function Invoke-Merge {
    if ($BundleFiles.Count -lt 2) { throw "-BundleFiles needs at least 2 bundles" }
    $all = @{}; $dirs = @{}; $n = 0
    foreach ($b in $BundleFiles) {
        $p = (Resolve-Path $b).Path; $n++
        Write-Host "Loading: $p"
        $data = Read-Bundle $p $Key
        foreach ($e in $data.Entries) {
            if ($all.ContainsKey($e.Path)) {
                $prev = $all[$e.Path]
                if ($prev.Hash -and $e.Hash -and $prev.Hash -ne $e.Hash) {
                    Write-Warning "conflict: $($e.Path)  (newer modification time wins)"
                    if ($e.MTime -ge $prev.MTime) { $all[$e.Path] = $e }
                }
            } else { $all[$e.Path] = $e }
        }
        foreach ($d in $data.Dirs) { $dirs[$d] = $true }
    }
    foreach ($k in @($dirs.Keys)) { if ($all.Keys | Where-Object { $_ -like "$k/*" }) { $dirs.Remove($k) } }   # not empty any more
    $dest = Get-AbsPath $BundleFile
    $tmp = "$dest.$([guid]::NewGuid().ToString('N').Substring(0,8)).sbtmp"
    $w = New-Object System.IO.StreamWriter($tmp, $false, $utf8); $w.NewLine = "`n"
    try {
        Write-BundleHeader $w "merged-$n-bundles" $all.Count $dirs.Count $Compress.IsPresent @($all.Values | Sort-Object Path | ForEach-Object { [pscustomobject]@{ Rel = $_.Path; Size = $_.Size } }) @($dirs.Keys | Sort-Object)
        foreach ($e in ($all.Values | Sort-Object Path)) { Write-BundleEntry $w $e }
        Write-BundleFooter $w @($dirs.Keys)
    } finally { $w.Dispose() }
    Complete-BundleFile $tmp $dest $Compress.IsPresent
    Write-Host ("OK: merged {0} bundle(s) -> {1} file(s), {2}" -f $n,$all.Count,(Format-Size (Get-Item $dest).Length)) -ForegroundColor Green
    Write-Host "  => $dest"
}

# ============================================================
# encrypt / decrypt
# ============================================================
function Invoke-Encrypt {
    if (-not $Key) { throw "-Key is required" }
    $p = (Get-BundleInput)
    $raw = [System.IO.File]::ReadAllBytes($p)
    if ([System.Text.Encoding]::ASCII.GetString($raw, 0, [Math]::Min(40, $raw.Length)).StartsWith('=== SOURCE-BUNDLE-ENCRYPTED')) { throw "already encrypted" }
    $out = "$p.enc"
    if ((Test-Path $out) -and -not $Overwrite) { throw "$out exists (use -Overwrite)" }
    $b64 = [Convert]::ToBase64String((Protect-Bytes $raw $Key))
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append($script:EncHeader).Append("`n")
    [void]$sb.Append("=== CIPHER: AES-256-CBC + HMAC-SHA256, PBKDF2 x$script:KdfIter ===`n`n")
    for ($i = 0; $i -lt $b64.Length; $i += 76) { [void]$sb.Append($b64.Substring($i, [Math]::Min(76, $b64.Length - $i))).Append("`n") }
    [System.IO.File]::WriteAllText($out, $sb.ToString(), $utf8)
    Write-Host "Encrypted: $out" -ForegroundColor Green
    Write-Host "  The original plain bundle is untouched - delete it yourself if you no longer need it." -ForegroundColor Yellow
}

function Invoke-Decrypt {
    if (-not $Key) { throw "-Key is required" }
    $p = (Get-BundleInput)
    $raw = [System.IO.File]::ReadAllBytes($p)
    if (-not [System.Text.Encoding]::ASCII.GetString($raw, 0, [Math]::Min(40, $raw.Length)).StartsWith('=== SOURCE-BUNDLE-ENCRYPTED')) { throw "not an encrypted bundle" }
    $lines = ($utf8.GetString($raw) -replace "`r", '') -split "`n"
    $b64 = ($lines | Where-Object { $_ -and -not $_.StartsWith('===') }) -join ''
    $plain = Unprotect-Bytes ([Convert]::FromBase64String($b64)) $Key
    $out = if ($p.EndsWith('.enc')) { $p.Substring(0, $p.Length - 4) } else { "$p.dec" }
    if ((Test-Path $out) -and -not $Overwrite) { throw "$out exists (use -Overwrite)" }
    [System.IO.File]::WriteAllBytes($out, $plain)
    Write-Host "Decrypted: $out" -ForegroundColor Green
}

function Invoke-GitCheck {
    $root = (Resolve-Path $Source).Path
    $gm = if ($GitIgnore -eq 'off') { 'auto' } else { $GitIgnore }
    $hasRepo = Test-Path (Join-Path $root '.git')
    $hasGit = [bool](Get-Command git -ErrorAction SilentlyContinue)
    Write-Host "Source        : $root" -ForegroundColor Cyan
    Write-Host ("git repository: {0}    git installed: {1}    -GitIgnore: {2}{3}" -f $hasRepo,$hasGit,$GitIgnore,$(if ($GitIgnore -eq 'off') { '  (checking as if auto)' } else { '' }))
    $flt = Get-IgnoreFilter $root $gm
    if (-not $flt) {
        Write-Host "no .gitignore / .bundleignore rules found - nothing is ignored (builtin excludes such as .git and node_modules still apply)" -ForegroundColor Yellow
    } else {
        Write-Host $flt.Info -ForegroundColor Green
        if ($flt.Kind -eq 'rules') {
            Write-Host "ignore files:" -ForegroundColor Yellow
            foreach ($g in ($flt.Set.Rules | Group-Object File)) { Write-Host ("  {0,-34} {1} rule(s)" -f $g.Name,$g.Count) }
        }
    }
    Write-Host "scanning..." -ForegroundColor DarkGray
    $scan = Get-SourceScan $root $flt $true
    $incB = [long]0; foreach ($f in $scan.Files) { $incB += $f.Size }
    $igFB = [long]0; foreach ($f in $scan.IgnoredFiles) { $igFB += $f.Size }
    $igDC = 0; $igDB = [long]0; foreach ($d in $scan.IgnoredDirs) { $igDC += $d.Count; $igDB += $d.Size }
    Write-Host ""
    Write-Host ("included : {0,6} files  {1}" -f $scan.Files.Count,(Format-Size $incB)) -ForegroundColor Green
    Write-Host ("ignored  : {0,6} files  {1}   (+ {2} folders: {3} files, {4})" -f $scan.IgnoredFiles.Count,(Format-Size $igFB),$scan.IgnoredDirs.Count,$igDC,(Format-Size $igDB)) -ForegroundColor Yellow
    if ($scan.IgnoredDirs.Count -gt 0) {
        Write-Host ""; Write-Host "ignored folders (largest first):" -ForegroundColor Yellow
        $scan.IgnoredDirs | Sort-Object Size -Descending | Select-Object -First 10 | ForEach-Object {
            Write-Host ("  {0,-36} {1,6} files  {2}" -f ($_.Rel + '/'),$_.Count,(Format-Size $_.Size)) }
    }
    if ($scan.SkippedDirs.Count -gt 0) {
        Write-Host ""; Write-Host "folders skipped by the default excludes (use -ForceInclude <name> or -NoDefaultExcludes to include):" -ForegroundColor Yellow
        $scan.SkippedDirs | Sort-Object Size -Descending | Select-Object -First 10 | ForEach-Object {
            Write-Host ("  {0,-36} {1,6} files  {2}" -f ($_.Rel + '/'),$_.Count,(Format-Size $_.Size)) }
    }
    if ($ForceInclude.Count -gt 0) { Write-Host ""; Write-Host ("force-included anyway: {0}" -f ($ForceInclude -join ', ')) -ForegroundColor Cyan }
    if ($scan.IgnoredFiles.Count -gt 0) {
        Write-Host ""; Write-Host "ignored files (largest first):" -ForegroundColor Yellow
        $scan.IgnoredFiles | Sort-Object Size -Descending | Select-Object -First 10 | ForEach-Object {
            Write-Host ("  {0,-48} {1}" -f $_.Rel,(Format-Size $_.Size)) }
    }
    $risky = @($scan.Files | Where-Object { $_.Rel -match '(^|/)(\.env(\..*)?|id_rsa|id_ed25519|.*\.pem|.*\.pfx|.*\.p12|credentials\.json)$' -and $_.Rel -notmatch '\.(example|sample|template)$' })
    if ($risky.Count -gt 0) {
        Write-Host ""; Write-Host "WARNING: sensitive-looking files are NOT ignored and will be packed:" -ForegroundColor Red
        $risky | ForEach-Object { Write-Host "  $($_.Rel)" -ForegroundColor Red }
        Write-Host "  fix: add them to .gitignore or .bundleignore (see ignore-suggest / secrets)" -ForegroundColor Red
    }
}

function Invoke-SrcTree {
    $root = (Resolve-Path $Source).Path
    $flt = Get-IgnoreFilter $root
    if ($flt) { Write-Host $flt.Info -ForegroundColor DarkGray } else { Write-Host "ignore files: not applied" -ForegroundColor DarkGray }
    $scan = Get-SourceScan $root $flt $true ''
    $items = @($scan.Files | ForEach-Object { [pscustomobject]@{ Rel = $_.Rel; Size = $_.Size } })
    $ed = if ($NoEmptyDirs) { @() } else { @($scan.EmptyDirs) }
    Write-PathTree $items $ed $Depth $true
    $ignN = $scan.IgnoredFiles.Count; foreach ($d in $scan.IgnoredDirs) { $ignN += $d.Count }
    Write-Host ""
    Write-Host ("{0} file(s) shown, {1} excluded by ignore files" -f $items.Count,$ignN) -ForegroundColor Green
    if ($Long -and $scan.IgnoredDirs.Count -gt 0) {
        Write-Host "ignored folders (not shown above):" -ForegroundColor Yellow
        $scan.IgnoredDirs | Sort-Object Rel | ForEach-Object { Write-Host ("  {0}/  [{1} files, {2}]" -f $_.Rel,$_.Count,(Format-Size $_.Size)) }
    }
}

function Invoke-BigFiles {
    $root = (Resolve-Path $Source).Path
    $flt = Get-IgnoreFilter $root
    if ($flt) { Write-Host $flt.Info -ForegroundColor DarkGray }
    $scan = Get-SourceScan $root $flt $false
    $n = if ($Top -gt 0) { $Top } else { 20 }
    Write-Host ("largest {0} files that WOULD be packed:" -f $n) -ForegroundColor Yellow
    $tot = [long]0; foreach ($f in $scan.Files) { $tot += $f.Size }
    $scan.Files | Sort-Object Size -Descending | Select-Object -First $n | ForEach-Object {
        $pct = if ($tot -gt 0) { [math]::Round(100 * $_.Size / $tot, 1) } else { 0 }
        $bin = if (Test-IsBinary $_.Full) { 'bin' } else { 'txt' }
        Write-Host ("  {0,10}  {1,5}%  {2}  {3}" -f (Format-Size $_.Size),$pct,$bin,$_.Rel)
    }
    Write-Host ""
    Write-Host ("total {0} files, {1}.  Exclude big ones with -Exclude / .bundleignore, or use -MaxTextKB / -MaxBinaryMB." -f $scan.Files.Count,(Format-Size $tot)) -ForegroundColor Green
}

function Invoke-Secrets {
    $root = (Resolve-Path $Source).Path
    $flt = Get-IgnoreFilter $root
    $scan = Get-SourceScan $root $flt $false
    $rules = $script:SecretRules
    $hits = 0; $filesHit = @{}
    Write-Host "scanning $($scan.Files.Count) file(s) that would be packed..." -ForegroundColor DarkGray
    foreach ($f in $scan.Files) {
        if ($f.Rel -match $script:SecretNameRx -and $f.Rel -notmatch '\.(example|sample|template)$') {
            Write-Host ("  {0}  [sensitive file name]" -f $f.Rel) -ForegroundColor Red; $hits++; $filesHit[$f.Rel] = 1
        }
        if ($f.Size -gt 1MB -or $f.Size -eq 0) { continue }
        if (Test-IsBinary $f.Full) { continue }
        $ln = 0
        foreach ($line in [System.IO.File]::ReadAllLines($f.Full)) {
            $ln++
            if ($line.Length -gt 2000) { continue }
            foreach ($r in $rules) {
                $m = [regex]::Match($line, $r.R)
                if ($m.Success) {
                    $v = $m.Value; $mask = $v.Substring(0, [Math]::Min(6, $v.Length)) + '****'
                    Write-Host ("  {0}:{1}  [{2}]  {3}" -f $f.Rel,$ln,$r.N,$mask) -ForegroundColor Yellow
                    $hits++; $filesHit[$f.Rel] = 1
                    break
                }
            }
            if ($hits -ge 300) { break }
        }
        if ($hits -ge 300) { Write-Host "  ... stopped after 300 findings" -ForegroundColor Red; break }
    }
    Write-Host ""
    if ($hits -eq 0) { Write-Host "no obvious secrets found (this is a heuristic scan, not a guarantee)" -ForegroundColor Green }
    else { Write-Host ("{0} potential secret(s) in {1} file(s). Remove them or ignore those files before sharing the bundle." -f $hits,$filesHit.Count) -ForegroundColor Red }
}

function Get-IgnoreSuggestions([string]$root) {
    $has = { param($pat) [bool](Get-ChildItem -Path $root -Filter $pat -File -Force -ErrorAction SilentlyContinue) }
    $s = @()
    $s += [pscustomobject]@{ Pattern = '.env';        Why = 'local secrets' }
    $s += [pscustomobject]@{ Pattern = '.env.*';      Why = 'local secrets (per environment)' }
    $s += [pscustomobject]@{ Pattern = '!.env.example'; Why = 'keep the template file' }
    $s += [pscustomobject]@{ Pattern = '*.log';       Why = 'log files' }
    $s += [pscustomobject]@{ Pattern = '.DS_Store';   Why = 'macOS metadata' }
    $s += [pscustomobject]@{ Pattern = 'Thumbs.db';   Why = 'Windows metadata' }
    $s += [pscustomobject]@{ Pattern = '*.swp';       Why = 'editor swap files' }
    $s += [pscustomobject]@{ Pattern = 'source-bundle*.txt'; Why = 'bundles made by this tool' }
    $s += [pscustomobject]@{ Pattern = '.source-bundle.json';  Why = 'this tool''s menu settings' }
    if (& $has 'package.json') {
        foreach ($p in 'node_modules/','dist/','.next/','.nuxt/','coverage/') { $s += [pscustomobject]@{ Pattern = $p; Why = 'Node project' } }
    }
    if ((& $has 'go.mod')) { foreach ($p in 'bin/','*.test','*.out') { $s += [pscustomobject]@{ Pattern = $p; Why = 'Go project' } } }
    if ((& $has 'requirements.txt') -or (& $has 'pyproject.toml') -or (& $has 'setup.py')) {
        foreach ($p in '__pycache__/','*.pyc','.venv/','venv/','.pytest_cache/','.mypy_cache/') { $s += [pscustomobject]@{ Pattern = $p; Why = 'Python project' } }
    }
    if ((& $has 'pom.xml') -or (& $has 'build.gradle') -or (& $has 'build.gradle.kts')) {
        foreach ($p in 'target/','build/','.gradle/') { $s += [pscustomobject]@{ Pattern = $p; Why = 'Java/Gradle project' } }
    }
    if ((& $has '*.csproj') -or (& $has '*.sln')) { foreach ($p in 'bin/','obj/','.vs/') { $s += [pscustomobject]@{ Pattern = $p; Why = '.NET project' } } }
    if (& $has 'Cargo.toml') { $s += [pscustomobject]@{ Pattern = 'target/'; Why = 'Rust project' } }
    if (& $has 'docker-compose*.yml') { $s += [pscustomobject]@{ Pattern = '*.override.yml'; Why = 'local compose overrides' } }
    # drop what the root .gitignore already lists
    $have = @{}
    $gi = Join-Path $root '.gitignore'
    if (Test-Path $gi -PathType Leaf) {
        foreach ($l in [System.IO.File]::ReadAllLines($gi)) { $have[$l.Trim().TrimStart('/').TrimEnd('/').ToLowerInvariant()] = $true }
    }
    $seen = @{}; $out = @()
    foreach ($x in $s) {
        $k = $x.Pattern.Trim().TrimStart('/').TrimEnd('/').ToLowerInvariant()
        if ($have.ContainsKey($k) -or $seen.ContainsKey($k)) { continue }
        $seen[$k] = $true; $out += $x
    }
    return $out
}

# prints suggestions; with -Overwrite appends the missing patterns to <Source>/.gitignore
function Invoke-Suggest {
    $root = (Resolve-Path $Source).Path
    $sug = @(Get-IgnoreSuggestions $root)
    if ($sug.Count -eq 0) { Write-Host ".gitignore already covers all common patterns for this project" -ForegroundColor Green; return }
    Write-Host "missing from .gitignore:" -ForegroundColor Yellow
    foreach ($x in $sug) { Write-Host ("  {0,-24} # {1}" -f $x.Pattern,$x.Why) }
    if ($Overwrite) {
        $gi = Join-Path $root '.gitignore'
        $txt = ''
        if ((Test-Path $gi -PathType Leaf) -and ([System.IO.File]::ReadAllText($gi).Length -gt 0) -and -not ([System.IO.File]::ReadAllText($gi).EndsWith("`n"))) { $txt += "`n" }
        $txt += "`n# added by Source-bundle`n" + (($sug | ForEach-Object { $_.Pattern }) -join "`n") + "`n"
        [System.IO.File]::AppendAllText($gi, $txt, $utf8)
        Write-Host ("appended {0} pattern(s) to {1}" -f $sug.Count,$gi) -ForegroundColor Green
    } else {
        Write-Host "(preview only - run with -Overwrite to append them to .gitignore)" -ForegroundColor DarkGray
    }
}

# ============================================================
# INTERACTIVE menu
#   * first screen: choose view -> [1] flat (all commands in one list)
#                                   [2] grouped (pick a group, then a command)
#   * both views share ONE command registry, so they never drift apart
#   * type a command name (or a unique prefix) anywhere to jump straight to it
#   * remembers your choices in .source-bundle.json (the key is never saved)
# NOTE: the Invoke-* functions read script-level variables ($BundleFile,
#       $Target, ...). The menu therefore sets $script:* before each call and
#       resets them afterwards (Reset-Ctx), so nothing leaks between commands.
#       Local variables in this section avoid those names on purpose.
# ============================================================
$script:SettingsFile = Join-Path (Get-Location).Path '.source-bundle.json'
$script:S = @{
    Bundle = $BundleFile; Source = '.'; Target = $Target; Menu = 'grouped'; Key = ''
    GitIgnore = $GitIgnore; IgnoreFiles = ($IgnoreFiles -join ','); SkipIgnore = $SkipIgnoreFiles.IsPresent
    NoDefault = $NoDefaultExcludes.IsPresent; NoEmpty = $NoEmptyDirs.IsPresent; Gzip = $Compress.IsPresent; Stamp = $false
    ExclSecrets = $ExcludeSecrets.IsPresent; NoIndex = $NoIndex.IsPresent; Clip = $Clipboard.IsPresent; ForceInclude = ($ForceInclude -join ',')
}
$script:Bound = @($PSBoundParameters.Keys)
$script:BoundMenu = $Menu
$script:LastCmd = $null
$script:Exit = $false

function Load-Settings {
    if (Test-Path $script:SettingsFile) {
        try {
            $j = Get-Content $script:SettingsFile -Raw | ConvertFrom-Json
            foreach ($k in 'Bundle','Source','Target','Menu','GitIgnore','IgnoreFiles','ForceInclude') {
                if ($j.PSObject.Properties[$k] -and $j.$k) { $script:S[$k] = [string]$j.$k }
            }
            foreach ($k in 'SkipIgnore','NoDefault','NoEmpty','Gzip','Stamp','ExclSecrets','NoIndex','Clip') {
                if ($j.PSObject.Properties[$k]) { $script:S[$k] = [bool]$j.$k }
            }
        } catch {}
    }
    # options typed on the command line win over remembered ones
    if ($script:Bound -contains 'BundleFile')         { $script:S.Bundle = $BundleFile }
    if ($script:Bound -contains 'Source')             { $script:S.Source = $Source }
    if ($script:Bound -contains 'Target')             { $script:S.Target = $Target }
    if ($script:Bound -contains 'GitIgnore')          { $script:S.GitIgnore = $GitIgnore }
    if ($script:Bound -contains 'IgnoreFiles')        { $script:S.IgnoreFiles = ($IgnoreFiles -join ',') }
    if ($script:Bound -contains 'SkipIgnoreFiles')    { $script:S.SkipIgnore = $SkipIgnoreFiles.IsPresent }
    if ($script:Bound -contains 'NoDefaultExcludes')  { $script:S.NoDefault = $NoDefaultExcludes.IsPresent }
    if ($script:Bound -contains 'NoEmptyDirs')        { $script:S.NoEmpty = $NoEmptyDirs.IsPresent }
    if ($script:Bound -contains 'Compress')           { $script:S.Gzip = $Compress.IsPresent }
    if ($script:Bound -contains 'ForceInclude')       { $script:S.ForceInclude = ($ForceInclude -join ',') }
    if ($script:Bound -contains 'IncludeAll' -and $IncludeAll) { $script:S.GitIgnore = 'off'; $script:S.NoDefault = $true }
    if ($script:Bound -contains 'ExcludeSecrets')     { $script:S.ExclSecrets = $ExcludeSecrets.IsPresent }
    if ($script:Bound -contains 'NoIndex')            { $script:S.NoIndex = $NoIndex.IsPresent }
    if ($script:Bound -contains 'Clipboard')          { $script:S.Clip = $Clipboard.IsPresent }
}

function Save-Settings {
    try {
        @{ Bundle = $script:S.Bundle; Source = $script:S.Source; Target = $script:S.Target; Menu = $script:S.Menu
           GitIgnore = $script:S.GitIgnore; IgnoreFiles = $script:S.IgnoreFiles; SkipIgnore = [bool]$script:S.SkipIgnore
           NoDefault = [bool]$script:S.NoDefault; NoEmpty = [bool]$script:S.NoEmpty; Gzip = [bool]$script:S.Gzip; Stamp = [bool]$script:S.Stamp; ForceInclude = $script:S.ForceInclude
           ExclSecrets = [bool]$script:S.ExclSecrets; NoIndex = [bool]$script:S.NoIndex; Clip = [bool]$script:S.Clip } |
            ConvertTo-Json | Set-Content -Path $script:SettingsFile -Encoding UTF8
    } catch {}
}

function Get-Abs([string]$p) {
    if ([string]::IsNullOrWhiteSpace($p)) { return $p }
    if ([System.IO.Path]::IsPathRooted($p)) { return $p }
    return [System.IO.Path]::GetFullPath((Join-Path (Get-Location).Path $p))
}
function Read-Default($prompt, $default) {
    $hint = if ($default -ne $null -and "$default" -ne '') { "[$default]" } else { '' }
    $v = Read-Host "  $prompt $hint"
    if ([string]::IsNullOrWhiteSpace($v)) { return $default }
    return $v.Trim()
}
function Read-YesNo($prompt, [bool]$default = $false) {
    $hint = if ($default) { 'Y/n' } else { 'y/N' }
    $v = Read-Host "  $prompt ($hint)"
    if ([string]::IsNullOrWhiteSpace($v)) { return $default }
    return ($v -match '^[yY]')
}
function Read-Int($prompt, [int]$default) {
    $v = Read-Host "  $prompt [$default]"
    $n = 0
    if ([string]::IsNullOrWhiteSpace($v)) { return $default }
    if ([int]::TryParse($v.Trim(), [ref]$n)) { return $n }
    Write-Host "  not a number, using $default" -ForegroundColor Yellow
    return $default
}
function Read-Double($prompt, [double]$default) {
    $v = Read-Host "  $prompt [$default]"
    $n = [double]0
    if ([string]::IsNullOrWhiteSpace($v)) { return $default }
    if ([double]::TryParse($v.Trim(), [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$n)) { return $n }
    Write-Host "  not a number, using $default" -ForegroundColor Yellow
    return $default
}
function Read-DateOrNone($prompt) {
    $v = Read-Host "  $prompt [none]"
    $d = [datetime]::MinValue
    if ([string]::IsNullOrWhiteSpace($v)) { return [datetime]::MinValue }
    if ([datetime]::TryParse($v.Trim(), [ref]$d)) { return $d }
    Write-Host "  could not parse date, ignoring" -ForegroundColor Yellow
    return [datetime]::MinValue
}
function Read-Secret([string]$prompt) {
    # piped / redirected input cannot use the masked prompt - fall back to plain text
    if ([Console]::IsInputRedirected) { return (Read-Host $prompt) }
    $ss = Read-Host $prompt -AsSecureString
    $ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($ss)
    try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr) }
}
function Split-List([string]$s) {
    return @($s -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}
function Pause-Menu { Write-Host ""; Read-Host "  Press Enter to continue" | Out-Null }

# Reset every script-level option the Invoke-* functions read.
# copy the remembered choices into the script-level options the Invoke-* functions read
function Sync-Ctx {
    $script:GitIgnore = [string]$script:S.GitIgnore
    $script:IgnoreFiles = @(Split-List ([string]$script:S.IgnoreFiles))
    $script:ForceInclude = @(Split-List ([string]$script:S.ForceInclude) | ForEach-Object { $_.Trim('/','\') })
    $script:SkipIgnoreFiles = [switch][bool]$script:S.SkipIgnore
    $script:NoDefaultExcludes = [switch][bool]$script:S.NoDefault
    $script:NoEmptyDirs = [switch][bool]$script:S.NoEmpty
    $script:Compress = [switch][bool]$script:S.Gzip
    $script:ExcludeSecrets = [switch][bool]$script:S.ExclSecrets
    $script:NoIndex = [switch][bool]$script:S.NoIndex
    $script:Clipboard = [switch][bool]$script:S.Clip
    $script:Key = [string]$script:S.Key
}
# reset every option, then apply the remembered choices
function Reset-Ctx {
    $script:Source = '.'; $script:Profile = ''; $script:Exclude = @(); $script:ExcludeDir = @(); $script:Include = @(); $script:OnlyExt = @()
    $script:NoBinary = [switch]$false; $script:MaxBinaryMB = 50; $script:MaxTextKB = 10240; $script:Since = [datetime]::MinValue
    $script:Target = [string]$script:S.Target; $script:DryRun = [switch]$false; $script:Overwrite = [switch]$false
    $script:Filter = '*'; $script:Path = ''; $script:DeleteExtras = [switch]$false; $script:PreserveTimes = [switch]$false
    $script:Pattern = ''; $script:CaseSensitive = [switch]$false; $script:Context = 0; $script:LineNumbers = [switch]$false
    $script:Long = [switch]$false; $script:Depth = 0; $script:CompareTo = '.'; $script:Compare2 = ''; $script:BundleFiles = @()
    $script:Top = 20; $script:IntervalSec = 3; $script:OutputJson = [switch]$false
    $script:SplitMB = 0; $script:Changed = [switch]$false; $script:ChangedSince = ''; $script:Base = ''; $script:FromClipboard = [switch]$false
    Sync-Ctx
    $script:BundleFile = Get-Abs $script:S.Bundle
}

function Run-Invoke([scriptblock]$sb) {
    for ($attempt = 0; $attempt -lt 2; $attempt++) {
        try { & $sb | Out-Host; return }
        catch {
            $msg = $_.Exception.Message
            if ($attempt -eq 0 -and $msg -match 'encrypted; supply -Key') {
                Write-Host "  this bundle is encrypted." -ForegroundColor Yellow
                $kk = Read-Secret '  key'
                if (-not $kk) { return }
                $script:Key = $kk; $script:S.Key = $kk
                continue
            }
            Write-Host "  ERROR: $msg" -ForegroundColor Red
            return
        }
    }
}

# Ask for a bundle file (defaults to the remembered one). Returns absolute path or $null.
function Pick-Bundle([string]$label = 'bundle file') {
    $cur = $script:S.Bundle
    $v = Read-Default $label $cur
    $abs = Get-Abs $v
    if (Test-Path $abs -PathType Leaf) { $script:S.Bundle = $v; return $abs }
    Write-Host "  not found: $abs" -ForegroundColor Yellow
    $cand = @(Get-ChildItem -Path (Get-Location).Path -File -Force -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like '*bundle*' -or $_.Name -like '*.enc' -or $_.Name -like 'merged*' } |
        Where-Object { $_.Name -notlike '*.sbtmp' } | Sort-Object LastWriteTime -Descending | Select-Object -First 9)
    if ($cand.Count -eq 0) { Write-Host "  no bundle files found in $((Get-Location).Path)" -ForegroundColor Yellow; return $null }
    Write-Host "  bundles found here:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $cand.Count; $i++) {
        Write-Host ("    [{0}] {1}  ({2}, {3:yyyy-MM-dd HH:mm})" -f ($i+1),$cand[$i].Name,(Format-Size $cand[$i].Length),$cand[$i].LastWriteTime)
    }
    $pick = Read-Host "  pick a number (Enter = cancel)"
    $idx = 0
    if ([int]::TryParse($pick, [ref]$idx) -and $idx -ge 1 -and $idx -le $cand.Count) {
        $script:S.Bundle = $cand[$idx-1].Name
        return $cand[$idx-1].FullName
    }
    return $null
}

function Get-BundleStatus {
    $bp = Get-Abs $script:S.Bundle
    if (Test-Path $bp -PathType Leaf) {
        $fi = Get-Item $bp
        $kind = ''
        try {
            $fs = [System.IO.File]::OpenRead($bp)
            try { $buf = New-Object byte[] 30; $n = $fs.Read($buf, 0, 30) } finally { $fs.Dispose() }
            if ($n -ge 2 -and $buf[0] -eq 0x1F -and $buf[1] -eq 0x8B) { $kind = ', gzip' }
            elseif ([System.Text.Encoding]::ASCII.GetString($buf, 0, $n).StartsWith('=== SOURCE-BUNDLE-ENCRYPTED')) { $kind = ', encrypted' }
        } catch {}
        return [pscustomobject]@{ Exists = $true; Text = ("{0}  ({1}{2}, {3:yyyy-MM-dd HH:mm})" -f $bp,(Format-Size $fi.Length),$kind,$fi.LastWriteTime) }
    }
    return [pscustomobject]@{ Exists = $false; Text = "$bp  (not created yet)" }
}

function Show-Header([string]$title) {
    Clear-Host
    Write-Host "==============================================================" -ForegroundColor Cyan
    Write-Host ("  Source-bundle v{0}    {1}" -f $script:AppVersion,$title) -ForegroundColor Cyan
    Write-Host "==============================================================" -ForegroundColor Cyan
    $st = Get-BundleStatus
    $col = if ($st.Exists) { 'Green' } else { 'Yellow' }
    Write-Host "  bundle : $($st.Text)" -ForegroundColor $col
    Write-Host ("  source : {0}    target : {1}    view : {2}    ignore files : {3}" -f $script:S.Source,$script:S.Target,$script:S.Menu,$script:S.GitIgnore) -ForegroundColor DarkGray
    if ($script:LastCmd) { Write-Host ("  last   : {0}   (type r to repeat)" -f $script:LastCmd.Name) -ForegroundColor DarkGray }
    Write-Host ""
}

# ============================================================
# command implementations (each one asks, sets $script:*, calls Invoke-*)
# ============================================================
# the choices that decide what ends up in the bundle - shown for every pack
# list what is being left out (ignored folders/files, default-skipped folders) and let the user pick what to include anyway
function Pick-ForceInclude([string]$src) {
    Write-Host "  scanning for what would be left out (this can take a moment) ..." -ForegroundColor DarkGray
    $root = (Resolve-Path $src).Path
    Sync-Ctx
    $keep = $script:ForceInclude; $script:ForceInclude = @()
    try {
        $flt = if ($script:S.GitIgnore -eq 'off') { $null } else { Get-IgnoreFilter $root 'auto' }
        $scan = Get-SourceScan $root $flt $true ''
    } finally { $script:ForceInclude = $keep }
    $cands = @()
    foreach ($d in $scan.IgnoredDirs) { $cands += [pscustomobject]@{ Rel = $d.Rel; Count = $d.Count; Size = $d.Size; Why = 'folder ignored by an ignore file' } }
    foreach ($d in $scan.SkippedDirs) { $cands += [pscustomobject]@{ Rel = $d.Rel; Count = $d.Count; Size = $d.Size; Why = 'folder skipped by default' } }
    $cands = @($cands | Sort-Object Size -Descending | Select-Object -First 18)
    $fl = @($scan.IgnoredFiles | Sort-Object Size -Descending | Select-Object -First 7)
    foreach ($f in $fl) { $cands += [pscustomobject]@{ Rel = $f.Rel; Count = 1; Size = $f.Size; Why = 'file ignored by an ignore file' } }
    $cur = @(Split-List ([string]$script:S.ForceInclude))
    if ($cands.Count -eq 0) { Write-Host "  nothing is being left out here" -ForegroundColor Green; return $cur }
    Write-Host ""
    for ($i = 0; $i -lt $cands.Count; $i++) {
        $c = $cands[$i]
        Write-Host ("    [{0,2}] {1,-36} {2,6} file(s) {3,9}   {4}" -f ($i + 1),$c.Rel,$c.Count,(Format-Size $c.Size),$c.Why)
    }
    $now = if ($cur.Count -gt 0) { $cur -join ', ' } else { 'none' }
    Write-Host "  currently included anyway: $now" -ForegroundColor DarkGray
    $a = Read-Host "  numbers to include anyway (comma separated, 'all', '-' = clear, Enter = keep current)"
    if ([string]::IsNullOrWhiteSpace($a)) { return $cur }
    if ($a.Trim() -eq '-') { return @() }
    $names = @()
    if ($a.Trim().ToLowerInvariant() -eq 'all') { $names = @($cands | ForEach-Object { $_.Rel }) }
    else {
        foreach ($tok in ($a -split '[,\s]+')) {
            $ix = 0
            if ([int]::TryParse($tok, [ref]$ix) -and $ix -ge 1 -and $ix -le $cands.Count) { $names += $cands[$ix - 1].Rel }
        }
    }
    return @($names | Select-Object -Unique)
}

# the choices that decide what ends up in the bundle - shown for every pack
function Ask-PackOptions([string]$src) {
    Write-Host ""
    Write-Host "  --- .gitignore and similar ignore files ---" -ForegroundColor Yellow
    $cands = '.gitignore','.dockerignore','.npmignore','.bundleignore','.eslintignore','.prettierignore','.gcloudignore','.vscodeignore'
    $found = @($cands | Where-Object { Test-Path (Join-Path $src $_) -PathType Leaf })
    if ($found.Count -gt 0) { Write-Host ("  found in the source folder: {0}" -f ($found -join ', ')) -ForegroundColor Green }
    else { Write-Host "  no ignore file in the source folder root (nested ones may still exist)" -ForegroundColor DarkGray }
    Write-Host "    [1] apply them          (built-in parser, nested ignore files too)"
    Write-Host "    [2] apply via real git  (git ls-files; needs git and a repository)"
    Write-Host "    [3] do not apply        (include everything, ignored files too)"
    Write-Host "    [4] apply them, but let me pick folders/files to include anyway  (e.g. node_modules)"
    $dflt = switch ($script:S.GitIgnore) { 'git' { '2' } 'off' { '3' } default { '1' } }
    $c = Read-Host "  choice [$dflt]"
    if ([string]::IsNullOrWhiteSpace($c)) { $c = $dflt }
    $pick = $false
    switch ($c.Trim()) { '2' { $script:S.GitIgnore = 'git' } '3' { $script:S.GitIgnore = 'off' } '4' { $script:S.GitIgnore = 'auto'; $pick = $true } default { $script:S.GitIgnore = 'auto' } }
    if ($script:S.GitIgnore -eq 'off') {
        $script:S.NoDefault = Read-YesNo 'also include node_modules / dist / build / .vscode ... (normally skipped)?' $true
    }
    if ($pick) {
        $list = Pick-ForceInclude $src
        $script:S.ForceInclude = ($list -join ',')
    }
    Write-Host ""
    Write-Host ("  ignore files honoured : {0}" -f $(if ($script:S.GitIgnore -eq 'off') { '(none - ignore files are not applied)' } else { $script:S.IgnoreFiles }))
    Write-Host ("  included anyway       : {0}" -f $(if ($script:S.ForceInclude) { $script:S.ForceInclude } else { '(nothing forced)' }))
    Write-Host ("  ignore files packed   : {0}" -f $(if ($script:S.SkipIgnore) { 'no  (.gitignore etc. left out)' } else { 'yes (.gitignore etc. are part of the bundle)' }))
    Write-Host ("  node_modules, dist..  : {0}" -f $(if ($script:S.NoDefault) { 'included' } else { 'skipped' }))
    Write-Host ("  empty folders         : {0}" -f $(if ($script:S.NoEmpty) { 'not recorded' } else { 'recorded' }))
    Write-Host ("  gzip                  : {0}" -f $(if ($script:S.Gzip) { 'yes (not readable as plain text)' } else { 'no (plain readable text)' }))
    Write-Host ("  file index at the top : {0}" -f $(if ($script:S.NoIndex) { 'no' } else { 'yes' }))
    Write-Host ("  leave out secrets     : {0}" -f $(if ($script:S.ExclSecrets) { 'yes (.env, keys, tokens ... are not packed)' } else { 'no' }))
    Write-Host ("  copy to clipboard     : {0}" -f $(if ($script:S.Clip) { 'yes' } else { 'no' }))
    if (Read-YesNo 'change these options?' $false) {
        if ($script:S.GitIgnore -ne 'off') {
            $script:S.IgnoreFiles = Read-Default 'ignore files to honour, comma separated' $script:S.IgnoreFiles
        }
        $fi = Read-Default "include anyway, comma separated (names like node_modules, or paths/patterns; '-' = clear)" $script:S.ForceInclude
        $script:S.ForceInclude = if ($fi -eq '-') { '' } else { $fi }
        $script:S.SkipIgnore = -not (Read-YesNo 'put the ignore files themselves (.gitignore ...) into the bundle?' (-not [bool]$script:S.SkipIgnore))
        $script:S.NoDefault  = Read-YesNo 'also include node_modules / dist / build / .vscode ... folders?' ([bool]$script:S.NoDefault)
        $script:S.NoEmpty    = -not (Read-YesNo 'record empty folders?' (-not [bool]$script:S.NoEmpty))
        $script:S.Gzip       = Read-YesNo 'gzip-compress the bundle?' ([bool]$script:S.Gzip)
        $script:S.NoIndex    = -not (Read-YesNo 'write a file index at the top of the bundle?' (-not [bool]$script:S.NoIndex))
        $script:S.ExclSecrets = Read-YesNo 'leave out files that look like secrets (.env, keys, tokens)?' ([bool]$script:S.ExclSecrets)
        $script:S.Clip       = Read-YesNo 'also copy the bundle text to the clipboard?' ([bool]$script:S.Clip)
    }
    Save-Settings; Sync-Ctx
}

function Get-OutPath([string]$name) {
    $abs = Get-Abs $name
    if (-not $script:S.Stamp) { return $abs }
    $dir = [System.IO.Path]::GetDirectoryName($abs); $leaf = [System.IO.Path]::GetFileNameWithoutExtension($abs); $ext = [System.IO.Path]::GetExtension($abs)
    return (Join-Path $dir ("{0}-{1:yyyyMMdd-HHmmss}{2}" -f $leaf,(Get-Date),$ext))
}
function Remember-Out([string]$outFile) { if (Test-Path $outFile -PathType Leaf) { $script:S.Bundle = $outFile } }

function Cmd-Pack([string]$prof, [string]$kind = '') {
    $src = Read-Default 'source folder' $script:S.Source
    if (-not (Test-Path $src -PathType Container)) { Write-Host "  source folder not found: $src" -ForegroundColor Red; return }
    $outDefault = if ($kind -eq 'incremental') { 'incremental.txt' } else { $script:S.Bundle }
    $baseFile = ''
    if ($kind -eq 'incremental') {
        $baseFile = Pick-Bundle 'base bundle (an earlier full pack)'; if (-not $baseFile) { return }
    }
    $out = Read-Default 'output bundle' $outDefault
    if ($kind -eq 'incremental' -and (Get-Abs $out) -eq $baseFile) { Write-Host "  output must differ from the base bundle" -ForegroundColor Red; return }
    $gitRef = ''; $partMb = 0
    if ($kind -eq 'changed') { $gitRef = Read-Default 'also include files changed since this git ref (Enter = only uncommitted changes)' '' }
    if ($kind -eq 'split')   { $partMb = Read-Double 'max size of each part in MB' 5 }
    Ask-PackOptions $src
    $script:S.Source = $src; if ($kind -eq '') { $script:S.Bundle = $out }; Save-Settings
    $script:Source = $src; $script:BundleFile = Get-OutPath $out; $script:Profile = $prof
    if ($kind -eq 'changed')     { $script:Changed = [switch]$true; $script:ChangedSince = $gitRef }
    if ($kind -eq 'incremental') { $script:Base = $baseFile }
    if ($kind -eq 'split')       { $script:SplitMB = [double]$partMb }
    if (Read-YesNo 'preview the file list first (nothing is written)?' $false) {
        $script:DryRun = [switch]$true
        Run-Invoke { Invoke-Pack }
        $script:DryRun = [switch]$false
        if (-not (Read-YesNo 'pack now?' $true)) { return }
    }
    Run-Invoke { Invoke-Pack }; if ($script:S.Stamp -and $kind -eq '') { Remember-Out $script:BundleFile }
}

function Cmd-PackAdvanced {
    $src = Read-Default 'source folder' $script:S.Source
    if (-not (Test-Path $src -PathType Container)) { Write-Host "  source folder not found: $src" -ForegroundColor Red; return }
    $out = Read-Default 'output bundle' $script:S.Bundle
    $pf  = (Read-Default 'start from profile (quick/code/docs/none)' 'none').ToLowerInvariant()
    $exc = Split-List (Read-Default 'exclude patterns, comma  (e.g. *.log,*.bak)' '')
    $exd = Split-List (Read-Default 'exclude folders, comma   (e.g. vendor,tmp)' '')
    $inc = Split-List (Read-Default 'include only these patterns, comma (e.g. src/*,*.md)' '')
    $ext = @(Split-List (Read-Default 'only extensions, comma (e.g. go,ts,md)' '') | ForEach-Object { $_.TrimStart('*').TrimStart('.').ToLowerInvariant() })
    $nobin = Read-YesNo 'skip binary files?' $false
    $mtk = Read-Int 'max text file size in KB' 10240
    $mbm = Read-Int 'max binary file size in MB' 50
    $sinceD = Read-DateOrNone 'only files modified since (e.g. 2026-11-01)'
    $partMb = Read-Double 'split into parts of N MB (0 = one file)' 0
    Ask-PackOptions $src
    $script:S.Source = $src; $script:S.Bundle = $out; Save-Settings
    $script:Source = $src; $script:BundleFile = Get-OutPath $out
    if (@('quick','code','docs') -contains $pf) { $script:Profile = $pf }
    $script:Exclude = $exc; $script:ExcludeDir = $exd; $script:Include = $inc; $script:OnlyExt = $ext
    $script:NoBinary = [switch]$nobin; $script:MaxTextKB = $mtk; $script:MaxBinaryMB = $mbm; $script:Since = $sinceD; $script:SplitMB = [double]$partMb
    Run-Invoke { Invoke-Pack }; if ($script:S.Stamp) { Remember-Out $script:BundleFile }
}

function Cmd-Watch {
    $src = Read-Default 'source folder' $script:S.Source
    if (-not (Test-Path $src -PathType Container)) { Write-Host "  source folder not found: $src" -ForegroundColor Red; return }
    $out = Read-Default 'output bundle' $script:S.Bundle
    $iv  = Read-Int 'check interval (seconds)' 3
    Ask-PackOptions $src
    Write-Host "  Ctrl+C stops watching (and exits the tool)." -ForegroundColor Yellow
    $script:S.Source = $src; $script:S.Bundle = $out; Save-Settings
    $script:Source = $src; $script:BundleFile = Get-Abs $out; $script:IntervalSec = $iv
    Run-Invoke { Invoke-Watch }
}

function Cmd-Unpack([string]$kind) {   # dry | direct | filter
    $bf = Pick-Bundle; if (-not $bf) { return }
    $tg = Read-Default 'target folder' $script:S.Target
    $script:S.Target = $tg
    $script:BundleFile = $bf; $script:Target = $tg
    if ($kind -eq 'dry') { $script:DryRun = [switch]$true }
    else {
        $script:Overwrite = [switch](Read-YesNo 'overwrite existing files that differ?' $false)
        $script:PreserveTimes = [switch](Read-YesNo 'restore original modified times?' $false)
    }
    if ($kind -eq 'filter') { $script:Filter = Read-Default 'filter on bundle path (wildcards, e.g. *.go or src/*)' '*.go' }
    Run-Invoke { Invoke-Unpack }
}

function Cmd-Paste {
    if (-not (Get-Command Get-Clipboard -ErrorAction SilentlyContinue)) { Write-Host "  Get-Clipboard is not available on this system" -ForegroundColor Red; return }
    Write-Host "  Copy a whole bundle (e.g. from a chat or a file) to the clipboard first." -ForegroundColor Yellow
    $tg = Read-Default 'target folder' $script:S.Target
    $script:S.Target = $tg; $script:Target = $tg; $script:FromClipboard = [switch]$true
    $script:DryRun = [switch](Read-YesNo 'preview only (dry run)?' $true)
    if (-not $script:DryRun.IsPresent) { $script:Overwrite = [switch](Read-YesNo 'overwrite existing files that differ?' $false) }
    Run-Invoke { Invoke-Unpack }
    Clear-TempFiles
}

function Cmd-Extract {
    $bf = Pick-Bundle; if (-not $bf) { return }
    $pp = Read-Default 'path inside bundle (file or folder, e.g. src/api)' ''
    if (-not $pp) { return }
    $tg = Read-Default 'target folder' 'extracted'
    $script:BundleFile = $bf; $script:Path = $pp; $script:Target = $tg
    $script:Overwrite = [switch](Read-YesNo 'overwrite existing files?' $true)
    Run-Invoke { Invoke-Extract }
}

function Cmd-Sync {
    $bf = Pick-Bundle; if (-not $bf) { return }
    $tg = Read-Default 'target folder' $script:S.Target
    $script:S.Target = $tg
    $script:BundleFile = $bf; $script:Target = $tg
    $script:DryRun = [switch](Read-YesNo 'preview only (dry run)?' $true)
    if (-not $script:DryRun.IsPresent) {
        $script:DeleteExtras = [switch](Read-YesNo 'DELETE files in target that are not in the bundle?' $false)
        if ($script:DeleteExtras.IsPresent -and -not (Read-YesNo "really delete extras in '$tg'?" $false)) { return }
    }
    Run-Invoke { Invoke-Sync }
}

function Cmd-Inspect([string]$what) {
    $bf = Pick-Bundle; if (-not $bf) { return }
    $script:BundleFile = $bf
    switch ($what) {
        'info'     { Run-Invoke { Invoke-Info } }
        'list'     { $script:OutputJson = [switch](Read-YesNo 'output as JSON?' $false); Run-Invoke { Invoke-List } }
        'listlong' { $script:Long = [switch]$true; Run-Invoke { Invoke-List } }
        'tree'     { $script:Depth = Read-Int 'depth (0 = all)' 0; Run-Invoke { Invoke-Tree } }
        'stats'    { $script:OutputJson = [switch](Read-YesNo 'output as JSON?' $false); Run-Invoke { Invoke-Stats } }
        'verify'   { Run-Invoke { Invoke-Verify } }
        'dups'     { Run-Invoke { Invoke-Duplicates } }
    }
}

function Cmd-Grep {
    $bf = Pick-Bundle; if (-not $bf) { return }
    $pt = Read-Default 'regex pattern' ''
    if (-not $pt) { return }
    $script:BundleFile = $bf; $script:Pattern = $pt
    $script:CaseSensitive = [switch](Read-YesNo 'case sensitive?' $false)
    $script:Context = Read-Int 'context lines around each hit' 0
    $script:LineNumbers = [switch](Read-YesNo 'show line numbers?' $true)
    Run-Invoke { Invoke-Grep }
}

function Cmd-Diff {
    $bf = Pick-Bundle; if (-not $bf) { return }
    $cmp = Read-Default 'compare against folder' $script:S.Source
    if (-not (Test-Path $cmp -PathType Container)) { Write-Host "  folder not found: $cmp" -ForegroundColor Red; return }
    $script:BundleFile = $bf; $script:CompareTo = $cmp
    Run-Invoke { Invoke-Diff }
}

function Cmd-Merge {
    $lst = Read-Default 'bundles to merge, comma separated (at least 2)' ''
    $items = Split-List $lst
    if ($items.Count -lt 2) { Write-Host "  need at least 2 bundles" -ForegroundColor Yellow; return }
    $abs = @($items | ForEach-Object { Get-Abs $_ })
    foreach ($a in $abs) { if (-not (Test-Path $a -PathType Leaf)) { Write-Host "  not found: $a" -ForegroundColor Red; return } }
    $out = Read-Default 'merged output file' 'merged.txt'
    $script:BundleFiles = $abs; $script:BundleFile = Get-Abs $out
    Run-Invoke { Invoke-Merge }
}

function Cmd-Encrypt {
    $bf = Pick-Bundle; if (-not $bf) { return }
    Write-Host "  AES-256 + HMAC. The encrypted copy is written next to the bundle as <bundle>.enc" -ForegroundColor Yellow
    Write-Host "  There is NO way to recover the bundle if you forget the key." -ForegroundColor Yellow
    $k1 = Read-Secret '  key'
    if (-not $k1) { return }
    $k2 = Read-Secret '  key again'
    if ($k1 -ne $k2) { Write-Host "  keys do not match - nothing changed" -ForegroundColor Red; return }
    $script:BundleFile = $bf; $script:Key = $k1
    if (Test-Path "$bf.enc") { $script:Overwrite = [switch](Read-YesNo "$bf.enc exists - overwrite?" $false) }
    Run-Invoke { Invoke-Encrypt }
}

function Cmd-Decrypt {
    $bf = Pick-Bundle 'encrypted bundle (.enc)'; if (-not $bf) { return }
    $k1 = Read-Secret '  key'
    if (-not $k1) { return }
    $script:BundleFile = $bf; $script:Key = $k1
    $plain = if ($bf.EndsWith('.enc')) { $bf.Substring(0, $bf.Length - 4) } else { "$bf.dec" }
    if (Test-Path $plain) { $script:Overwrite = [switch](Read-YesNo "$plain exists - overwrite?" $false) }
    Run-Invoke { Invoke-Decrypt }
}

function Cmd-Schema { Run-Invoke { Write-Schema } }


function Cmd-GitCheck {
    $src = Read-Default 'source folder' $script:S.Source
    if (-not (Test-Path $src -PathType Container)) { Write-Host "  source folder not found: $src" -ForegroundColor Red; return }
    $script:S.Source = $src; $script:Source = $src
    Run-Invoke { Invoke-GitCheck }
}
function Cmd-SrcTree {
    $src = Read-Default 'source folder' $script:S.Source
    if (-not (Test-Path $src -PathType Container)) { Write-Host "  source folder not found: $src" -ForegroundColor Red; return }
    $script:S.Source = $src; $script:Source = $src
    $script:Depth = Read-Int 'depth (0 = all)' 0
    $script:Long = [switch](Read-YesNo 'also list the ignored folders?' $true)
    Run-Invoke { Invoke-SrcTree }
}
function Cmd-BigFiles {
    $src = Read-Default 'source folder' $script:S.Source
    if (-not (Test-Path $src -PathType Container)) { Write-Host "  source folder not found: $src" -ForegroundColor Red; return }
    $script:S.Source = $src; $script:Source = $src
    $script:Top = Read-Int 'how many files to show' 20
    Run-Invoke { Invoke-BigFiles }
}
function Cmd-Secrets {
    $src = Read-Default 'source folder' $script:S.Source
    if (-not (Test-Path $src -PathType Container)) { Write-Host "  source folder not found: $src" -ForegroundColor Red; return }
    $script:S.Source = $src; $script:Source = $src
    Run-Invoke { Invoke-Secrets }
}
function Cmd-Suggest {
    $src = Read-Default 'source folder' $script:S.Source
    if (-not (Test-Path $src -PathType Container)) { Write-Host "  source folder not found: $src" -ForegroundColor Red; return }
    $script:S.Source = $src; $script:Source = $src
    Run-Invoke { Invoke-Suggest }
    $missing = @(Get-IgnoreSuggestions (Resolve-Path $src).Path)
    if ($missing.Count -gt 0 -and (Read-YesNo 'append the missing patterns to .gitignore?' $false)) {
        $script:Overwrite = [switch]$true
        Run-Invoke { Invoke-Suggest }
    }
}
function Cmd-DiffBundles {
    $keep = $script:S.Bundle
    $b1 = Pick-Bundle 'first bundle (older)'; if (-not $b1) { return }
    $b2 = Pick-Bundle 'second bundle (newer)'; $script:S.Bundle = $keep
    if (-not $b2) { return }
    $script:BundleFile = $b1; $script:Compare2 = $b2
    Run-Invoke { Invoke-BundleDiff }
}

# ============================================================
# command registry (single source of truth for both views)
# ============================================================
function New-Cmd([string]$Name,[string]$Group,[string]$Desc,[scriptblock]$Act,[bool]$NoPause = $false) {
    [pscustomobject]@{ Num = 0; Name = $Name; Group = $Group; Desc = $Desc; Act = $Act; NoPause = $NoPause }
}

$script:Groups = @(
    [pscustomobject]@{ Key='pack';    Title='PACK';              Desc='project folder -> one text file' }
    [pscustomobject]@{ Key='restore'; Title='RESTORE';           Desc='rebuild folders + files from the bundle' }
    [pscustomobject]@{ Key='inspect'; Title='INSPECT';           Desc='info / list / tree / stats / verify' }
    [pscustomobject]@{ Key='search';  Title='SEARCH & COMPARE';  Desc='grep / diff / duplicates / merge' }
    [pscustomobject]@{ Key='git';     Title='IGNORE FILES & SAFETY'; Desc='.gitignore check / source tree / big files / secrets' }
    [pscustomobject]@{ Key='secure';  Title='ENCRYPT';           Desc='AES-256 encrypt / decrypt a bundle' }
)

$script:Cmds = @(
    New-Cmd 'pack'          'pack'    'whole project, every file (your ignore choices)' { Cmd-Pack '' }
    New-Cmd 'quick'         'pack'    'no logs/backups, binary files skipped'      { Cmd-Pack 'quick' }
    New-Cmd 'code'          'pack'    'source-code and config files only'          { Cmd-Pack 'code' }
    New-Cmd 'docs'          'pack'    'documentation files only'                   { Cmd-Pack 'docs' }
    New-Cmd 'changed'       'pack'    'only files git reports as changed'          { Cmd-Pack '' 'changed' }
    New-Cmd 'incremental'   'pack'    'only files that differ from an earlier bundle' { Cmd-Pack '' 'incremental' }
    New-Cmd 'split'         'pack'    'several parts of N MB (for chat size limits)' { Cmd-Pack '' 'split' }
    New-Cmd 'advanced'      'pack'    'custom filters, sizes, dates'               { Cmd-PackAdvanced }
    New-Cmd 'watch'         'pack'    're-pack automatically on file changes'      { Cmd-Watch } $true
    New-Cmd 'dry-run'       'restore' 'preview what unpack would write'            { Cmd-Unpack 'dry' }
    New-Cmd 'unpack'        'restore' 'rebuild the whole project in a folder'      { Cmd-Unpack 'direct' }
    New-Cmd 'unpack-filter' 'restore' 'rebuild only paths matching a wildcard'     { Cmd-Unpack 'filter' }
    New-Cmd 'paste'         'restore' 'rebuild the project from the clipboard'     { Cmd-Paste }
    New-Cmd 'extract'       'restore' 'restore one file or sub-folder'             { Cmd-Extract }
    New-Cmd 'sync'          'restore' 'update a folder, overwrite only if changed' { Cmd-Sync }
    New-Cmd 'info'          'inspect' 'bundle header + counts'                     { Cmd-Inspect 'info' }
    New-Cmd 'list'          'inspect' 'list files'                                 { Cmd-Inspect 'list' }
    New-Cmd 'list-long'     'inspect' 'list with mtime / hash / line endings'      { Cmd-Inspect 'listlong' }
    New-Cmd 'tree'          'inspect' 'show the bundle as a tree'                  { Cmd-Inspect 'tree' }
    New-Cmd 'stats'         'inspect' 'sizes by extension / top folder'            { Cmd-Inspect 'stats' }
    New-Cmd 'verify'        'inspect' 'check sha256 of every file'                 { Cmd-Inspect 'verify' }
    New-Cmd 'schema'        'inspect' 'print the bundle format'                    { Cmd-Schema }
    New-Cmd 'grep'          'search'  'regex search inside the bundle'             { Cmd-Grep }
    New-Cmd 'diff'          'search'  'compare bundle against a live folder'       { Cmd-Diff }
    New-Cmd 'diff-bundles'  'search'  'compare two bundles'                        { Cmd-DiffBundles }
    New-Cmd 'duplicates'    'search'  'find files with identical content'          { Cmd-Inspect 'dups' }
    New-Cmd 'merge'         'search'  'combine several bundles into one'           { Cmd-Merge }
    New-Cmd 'gitcheck'      'git'     'what the ignore files exclude + risky files' { Cmd-GitCheck }
    New-Cmd 'src-tree'      'git'     'live folder tree, ignore rules applied'     { Cmd-SrcTree }
    New-Cmd 'bigfiles'      'git'     'largest files that would be packed'         { Cmd-BigFiles }
    New-Cmd 'secrets'       'git'     'scan for keys / tokens / passwords'         { Cmd-Secrets }
    New-Cmd 'ignore-suggest' 'git'    'suggest missing .gitignore patterns'        { Cmd-Suggest }
    New-Cmd 'encrypt'       'secure'  'encrypt a bundle with a key (.enc)'         { Cmd-Encrypt }
    New-Cmd 'decrypt'       'secure'  'decrypt a .enc bundle'                      { Cmd-Decrypt }
)
$n_ = 0; foreach ($c_ in $script:Cmds) { $n_++; $c_.Num = $n_ }

# ============================================================
# input resolution + running
# ============================================================
function Find-ByName([string]$t) {
    $t = $t.Trim().ToLowerInvariant()
    $cx = @($script:Cmds | Where-Object { $_.Name -eq $t })
    if ($cx.Count -eq 1) { return @{ Type = 'cmd'; Cmd = $cx[0] } }
    $gx = @($script:Groups | Where-Object { $_.Key -eq $t })
    if ($gx.Count -eq 1) { return @{ Type = 'group'; Group = $gx[0] } }
    $cp = @($script:Cmds   | Where-Object { $_.Name -like "$t*" })
    $gp = @($script:Groups | Where-Object { $_.Key  -like "$t*" })
    if (($cp.Count + $gp.Count) -eq 1) {
        if ($cp.Count -eq 1) { return @{ Type = 'cmd'; Cmd = $cp[0] } }
        return @{ Type = 'group'; Group = $gp[0] }
    }
    if (($cp.Count + $gp.Count) -gt 1) {
        $names = @($cp | ForEach-Object { $_.Name }) + @($gp | ForEach-Object { $_.Key })
        return @{ Type = 'ambiguous'; Names = ($names -join ', ') }
    }
    return @{ Type = 'none' }
}

# Keys that work on every screen. Returns $true if handled.
function Handle-Nav([string]$c) {
    switch -Regex ($c.Trim().ToLowerInvariant()) {
        '^(q|quit|exit)$' { $script:Exit = $true; return $true }
        '^m$' {
            if ($script:S.Menu -eq 'flat') { $script:S.Menu = 'grouped' } else { $script:S.Menu = 'flat' }
            Save-Settings; $script:SwitchView = $true; return $true
        }
        '^s$' { Show-Settings; return $true }
        '^(\?|h|help)$' { Show-MenuHelp; return $true }
        '^r$' {
            if ($script:LastCmd) { Invoke-CmdObj $script:LastCmd } else { Write-Host "  nothing to repeat yet" -ForegroundColor Yellow; Start-Sleep 1 }
            return $true
        }
    }
    return $false
}

function Invoke-CmdObj($cmd) {
    $script:LastCmd = $cmd
    Reset-Ctx
    Write-Host ""
    Write-Host ("--- {0} ---" -f $cmd.Name.ToUpper()) -ForegroundColor Cyan
    try { & $cmd.Act } catch { Write-Host "  ERROR: $($_.Exception.Message)" -ForegroundColor Red }
    Clear-TempFiles
    Save-Settings
    if (-not $cmd.NoPause) { Pause-Menu }
}

function Run-ByName([string]$c) {
    $r = Find-ByName $c
    switch ($r.Type) {
        'cmd'       { Invoke-CmdObj $r.Cmd }
        'group'     { Show-Group $r.Group }
        'ambiguous' { Write-Host "  ambiguous: $($r.Names)" -ForegroundColor Yellow; Start-Sleep -Seconds 2 }
        default     { Write-Host "  unknown choice: $c" -ForegroundColor Red; Start-Sleep -Seconds 1 }
    }
}

# ============================================================
# screens
# ============================================================
function Show-MenuHelp {
    Show-Header 'HELP'
    Write-Host "  Keys (work on every screen):" -ForegroundColor Yellow
    Write-Host "    number        run that command / open that group"
    Write-Host "    name          type a command or group name, or a unique prefix (e.g. gre, tr, ins)"
    Write-Host "    /text         flat view: show only commands matching text"
    Write-Host "    m             switch between flat and grouped view"
    Write-Host "    r             repeat the last command"
    Write-Host "    s             settings (bundle / source / target / key / ignore-file choices / timestamps)"
    Write-Host "    0 or b        back          q  quit"
    Write-Host ""
    Write-Host "  Tip: every prompt shows its default in [brackets] - press Enter to accept." -ForegroundColor DarkGray
    Write-Host "  Your choices are remembered in .source-bundle.json (the key is never saved)." -ForegroundColor DarkGray
    Write-Host ""
    if (Read-YesNo 'show full command-line help too?' $false) { Write-Help }
    Pause-Menu
}

function Show-Settings {
    while ($true) {
        Show-Header 'SETTINGS'
        $yn = { param($v) if ($v) { 'yes' } else { 'no' } }
        Write-Host ("  [1] default bundle file   : {0}" -f $script:S.Bundle)
        Write-Host ("  [2] default source folder : {0}" -f $script:S.Source)
        Write-Host ("  [3] default target folder : {0}" -f $script:S.Target)
        $kt = if ($script:S.Key) { 'set (this session only)' } else { 'not set' }
        Write-Host ("  [4] session key           : {0}" -f $kt)
        Write-Host ("  [5] default view          : {0}" -f $script:S.Menu)
        Write-Host ""
        Write-Host "  what goes into a bundle:" -ForegroundColor Yellow
        Write-Host ("  [6] ignore files          : {0}   (auto = built-in parser | git = real git | off = apply none)" -f $script:S.GitIgnore)
        Write-Host ("  [7] ignore files honoured : {0}" -f $script:S.IgnoreFiles)
        Write-Host ("  [8] pack the ignore files themselves (.gitignore ...) : {0}" -f (& $yn (-not $script:S.SkipIgnore)))
        Write-Host ("  [9] include node_modules, dist, build ...             : {0}" -f (& $yn $script:S.NoDefault))
        Write-Host ("  [a] record empty folders                              : {0}" -f (& $yn (-not $script:S.NoEmpty)))
        Write-Host ("  [b] gzip-compress bundles                             : {0}" -f (& $yn $script:S.Gzip))
        Write-Host ("  [c] timestamped bundle names                          : {0}" -f (& $yn $script:S.Stamp))
        Write-Host ("  [d] leave out secrets (.env, keys, tokens)            : {0}" -f (& $yn $script:S.ExclSecrets))
        Write-Host ("  [e] file index at the top of bundles                  : {0}" -f (& $yn (-not $script:S.NoIndex)))
        Write-Host ("  [f] copy bundle to the clipboard after packing        : {0}" -f (& $yn $script:S.Clip))
        Write-Host ("  [g] always include (even if ignored / default-skipped): {0}" -f $(if ($script:S.ForceInclude) { $script:S.ForceInclude } else { '(nothing forced)' }))
        Write-Host "  [0] back"
        Write-Host ""
        $c = Read-Host "  choice"
        switch ($c) {
            '1' { $script:S.Bundle = Read-Default 'bundle file' $script:S.Bundle }
            '2' { $script:S.Source = Read-Default 'source folder' $script:S.Source }
            '3' { $script:S.Target = Read-Default 'target folder' $script:S.Target }
            '4' {
                if ($script:S.Key) { if (Read-YesNo 'clear the session key?' $true) { $script:S.Key = '' } }
                else { $script:S.Key = Read-Secret '  key' }
            }
            '5' { if ($script:S.Menu -eq 'flat') { $script:S.Menu = 'grouped' } else { $script:S.Menu = 'flat' } }
            '6' { $order = @('auto','git','off'); $ix = [array]::IndexOf($order, $script:S.GitIgnore); $script:S.GitIgnore = $order[($ix + 1) % $order.Count] }
            '7' { $script:S.IgnoreFiles = Read-Default 'ignore files, comma separated (e.g. .gitignore,.dockerignore,.bundleignore)' $script:S.IgnoreFiles }
            '8' { $script:S.SkipIgnore = -not [bool]$script:S.SkipIgnore }
            '9' { $script:S.NoDefault = -not [bool]$script:S.NoDefault }
            'a' { $script:S.NoEmpty = -not [bool]$script:S.NoEmpty }
            'b' { $script:S.Gzip = -not [bool]$script:S.Gzip }
            'c' { $script:S.Stamp = -not [bool]$script:S.Stamp }
            'd' { $script:S.ExclSecrets = -not [bool]$script:S.ExclSecrets }
            'e' { $script:S.NoIndex = -not [bool]$script:S.NoIndex }
            'f' { $script:S.Clip = -not [bool]$script:S.Clip }
            'g' { $fi = Read-Default "names/paths/patterns, comma separated ('-' = clear)" $script:S.ForceInclude; $script:S.ForceInclude = if ($fi -eq '-') { '' } else { $fi } }
            { $_ -in '0','x','' } { Save-Settings; return }
        }
        Save-Settings
    }
}

function Show-Group($g) {
    $items = @($script:Cmds | Where-Object { $_.Group -eq $g.Key })
    while (-not $script:Exit) {
        Show-Header $g.Title
        for ($i = 0; $i -lt $items.Count; $i++) {
            Write-Host ("  [{0}] {1,-14} {2}" -f ($i+1),$items[$i].Name,$items[$i].Desc)
        }
        Write-Host ""
        Write-Host "  [0] back    [m] view    [s] settings    [r] repeat    [?] help    [q] quit" -ForegroundColor DarkGray
        $c = Read-Host "  choice"
        if ($null -eq $c) { $script:Exit = $true; return }
        $c = $c.Trim()
        if ($c -eq '') { continue }
        if ($c -in '0','b','back') { return }
        $script:SwitchView = $false
        if (Handle-Nav $c) { if ($script:SwitchView) { return }; continue }
        $num = 0
        if ([int]::TryParse($c, [ref]$num)) {
            if ($num -ge 1 -and $num -le $items.Count) { Invoke-CmdObj $items[$num-1] }
            else { Write-Host "  no such number" -ForegroundColor Red; Start-Sleep 1 }
            continue
        }
        Run-ByName $c
    }
}

function Show-Grouped {
    while (-not $script:Exit -and $script:S.Menu -eq 'grouped') {
        Show-Header 'GROUPED VIEW'
        for ($i = 0; $i -lt $script:Groups.Count; $i++) {
            $g = $script:Groups[$i]
            $cnt = @($script:Cmds | Where-Object { $_.Group -eq $g.Key }).Count
            Write-Host ("  [{0}] {1,-22} {2}  ({3} commands)" -f ($i+1),$g.Title,$g.Desc,$cnt)
        }
        Write-Host ""
        Write-Host "  [m] flat view    [s] settings    [r] repeat    [?] help    [q] quit" -ForegroundColor DarkGray
        Write-Host "  Tip: you can also type a command name directly, e.g. grep" -ForegroundColor DarkGray
        $c = Read-Host "  choose a group"
        if ($null -eq $c) { $script:Exit = $true; return }
        $c = $c.Trim()
        if ($c -eq '') { continue }
        $script:SwitchView = $false
        if (Handle-Nav $c) { continue }
        $num = 0
        if ([int]::TryParse($c, [ref]$num)) {
            if ($num -ge 1 -and $num -le $script:Groups.Count) { Show-Group $script:Groups[$num-1] }
            else { Write-Host "  no such group" -ForegroundColor Red; Start-Sleep 1 }
            continue
        }
        Run-ByName $c
    }
}

function Show-Flat {
    $ftxt = ''
    while (-not $script:Exit -and $script:S.Menu -eq 'flat') {
        $title = if ($ftxt) { "FLAT VIEW  /$ftxt" } else { 'FLAT VIEW' }
        Show-Header $title
        $shown = 0
        foreach ($g in $script:Groups) {
            $rows = @($script:Cmds | Where-Object {
                $_.Group -eq $g.Key -and (-not $ftxt -or $_.Name -like "*$ftxt*" -or $_.Desc -like "*$ftxt*") })
            if ($rows.Count -eq 0) { continue }
            Write-Host ("  -- {0} --" -f $g.Title) -ForegroundColor DarkCyan
            foreach ($r in $rows) { Write-Host ("  [{0,2}] {1,-14} {2}" -f $r.Num,$r.Name,$r.Desc); $shown++ }
        }
        if ($shown -eq 0) { Write-Host "  no command matches '$ftxt'" -ForegroundColor Yellow }
        Write-Host ""
        Write-Host "  [/text] filter    [m] grouped view    [s] settings    [r] repeat    [?] help    [q] quit" -ForegroundColor DarkGray
        $c = Read-Host "  choice"
        if ($null -eq $c) { $script:Exit = $true; return }
        $c = $c.Trim()
        if ($c -eq '') { $ftxt = ''; continue }
        if ($c.StartsWith('/')) { $ftxt = $c.Substring(1).Trim(); continue }
        $script:SwitchView = $false
        if (Handle-Nav $c) { continue }
        $num = 0
        if ([int]::TryParse($c, [ref]$num)) {
            $hit = @($script:Cmds | Where-Object { $_.Num -eq $num })
            if ($hit.Count -eq 1) { Invoke-CmdObj $hit[0] }
            else { Write-Host "  no such number" -ForegroundColor Red; Start-Sleep 1 }
            continue
        }
        Run-ByName $c
    }
}

function Select-ViewMode {
    Show-Header 'START'
    $dflt = if ($script:S.Menu -eq 'flat') { '1' } else { '2' }
    Write-Host "  How do you want to browse the commands?"
    Write-Host ""
    Write-Host ("  [1] flat      all {0} commands in one numbered list" -f $script:Cmds.Count)
    Write-Host ("  [2] grouped   pick a group first ({0} groups), then a command" -f $script:Groups.Count)
    Write-Host "  [q] quit"
    Write-Host ""
    $c = Read-Host "  choice [$dflt]"
    if ($null -eq $c) { return 'quit' }
    if ([string]::IsNullOrWhiteSpace($c)) { $c = $dflt }
    switch -Regex ($c.Trim().ToLowerInvariant()) {
        '^(1|f|flat)$'    { return 'flat' }
        '^(2|g|grouped)$' { return 'grouped' }
        '^(q|quit|exit)$' { return 'quit' }
    }
    return $script:S.Menu
}

function Invoke-Interactive {
    Load-Settings
    $view = $script:BoundMenu
    if ($view -eq 'ask') { $view = Select-ViewMode }
    if ($view -eq 'quit') { Write-Host "bye" -ForegroundColor Green; return }
    $script:S.Menu = $view
    Save-Settings
    while (-not $script:Exit) {
        if ($script:S.Menu -eq 'flat') { Show-Flat } else { Show-Grouped }
    }
    Save-Settings
    Write-Host "bye" -ForegroundColor Green
}

# ============================================================
# dispatch
# ============================================================
switch ($Mode) {
    'interactive' { Invoke-Interactive }
    'pack'        { Invoke-Pack }
    'unpack'      { Invoke-Unpack }
    'info'        { Invoke-Info }
    'list'        { Invoke-List }
    'tree'        { Invoke-Tree }
    'extract'     { Invoke-Extract }
    'grep'        { Invoke-Grep }
    'verify'      { Invoke-Verify }
    'stats'       { Invoke-Stats }
    'diff'        { Invoke-Diff }
    'bdiff'       { Invoke-BundleDiff }
    'merge'       { Invoke-Merge }
    'sync'        { Invoke-Sync }
    'watch'       { Invoke-Watch }
    'encrypt'     { Invoke-Encrypt }
    'decrypt'     { Invoke-Decrypt }
    'duplicates'  { Invoke-Duplicates }
    'gitcheck'    { Invoke-GitCheck }
    'srctree'     { Invoke-SrcTree }
    'bigfiles'    { Invoke-BigFiles }
    'secrets'     { Invoke-Secrets }
    'suggest'     { Invoke-Suggest }
    'schema'      { Write-Schema }
    'help'        { Write-Help }
    default       { Write-Help }
}
Clear-TempFiles
if ($script:VerifyFailed -and $Mode -ne 'interactive') { exit 1 }
