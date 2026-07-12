<#
.SYNOPSIS
    長期間更新されていないファイルを robocopy で高速に列挙し、経過年数ごとに CSV へ出力します。

.DESCRIPTION
    Get-ChildItem による再帰列挙は、ファイル数が多い環境では処理時間が問題になります。
    本スクリプトは robocopy のリストモード (/L) を列挙エンジンとして利用し、
    さらに複数フォルダを並列に走査することで、走査時間を短縮します。

    出力は経過年数ごとに 3 つの CSV に分割します (重複なし)。
      StaleFiles_1y_*.csv : 1年以上 3年未満 更新なし
      StaleFiles_3y_*.csv : 3年以上 5年未満 更新なし
      StaleFiles_5y_*.csv : 5年以上       更新なし

    3 つの CSV は処理が終わるまで開いたまま、並行して書き込みます。
    あわせて、トップフォルダ単位の集計ログを出力します。

    本ツールはファイルの削除・移動を一切行いません。

    走査対象・出力先・並列数・出力ファイル名・バッファサイズは、
    スクリプト冒頭の「設定」ブロックで変更できます。

.PARAMETER Path
    走査対象のルートフォルダ。カンマ区切りで複数指定できます。
    優先順位: この引数 > 設定 $DEFAULT_TARGET_PATHS > ユーザープロファイル配下

.PARAMETER OutputDir
    出力先フォルダ。
    優先順位: この引数 > 設定 $DEFAULT_OUTPUT_DIR > スクリプトが置かれているフォルダ

.EXAMPLE
    .\Find-StaleFiles.ps1

.EXAMPLE
    .\Find-StaleFiles.ps1 -Path "D:\Data" -OutputDir "D:\Report"

.EXAMPLE
    .\Find-StaleFiles.ps1 -Path "D:\資材", "F:\"
#>

[CmdletBinding()]
param(
    [string[]] $Path,
    [string]   $OutputDir
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ############################################################
#
#  設定 (ここを編集すれば動作を変更できます)
#
# ############################################################

# ---- 走査対象フォルダ (複数指定可) --------------------------
#   下の @( ) の中に、走査したいフォルダを列挙してください。
#   行頭の # を外すと有効になります。カンマは不要です。
#   空のままにすると、ユーザープロファイル配下を走査します。
#
#   優先順位: 引数 -Path > この設定 > ユーザープロファイル
#
#   ↓↓↓ ここを編集してください ↓↓↓
$DEFAULT_TARGET_PATHS = @(

    # 'F:\'
    # 'D:\'
    # 'C:\Users\PC_User\OneDrive\'

)
#   ↑↑↑ ここまで ↑↑↑

# ---- 出力先フォルダ ----------------------------------------
#   空文字 '' のままにすると、スクリプトが置かれているフォルダへ出力します。
#     書き方の例:  'D:\Report'
$DEFAULT_OUTPUT_DIR = ''

# ---- 並列数 ------------------------------------------------
#   同時に走査するフォルダ数の上限。
#   増やせば速くなるとは限らない。ディスクの読み取りが上限に達すると、
#   むしろ待ち時間が増えるため、既定は 3 とする。
$MAX_PARALLEL = 3

# ---- 出力ファイル名 ----------------------------------------
#   {0} には実行日時 (yyyyMMdd_HHmmss) が入ります。
#   {0} を削除すると、実行のたびに同じ名前で上書きされます。
$CSV_NAME_1Y = 'StaleFiles_1y_{0}.csv'
$CSV_NAME_3Y = 'StaleFiles_3y_{0}.csv'
$CSV_NAME_5Y = 'StaleFiles_5y_{0}.csv'
$LOG_NAME    = 'StaleFiles_summary_{0}.log'

# ---- バッファサイズ ----------------------------------------
#   1 行ごとにディスクへ書き出すと I/O が処理のボトルネックになるため、
#   バッファに溜めてからまとめて書き出す。
$BUFFER_SIZE = 8KB

# ---- CSV のヘッダ ------------------------------------------
$CSV_HEADER = 'ファイル名,拡張子,サイズ(KB),最終更新日時,経過日数,階層,フォルダパス'

# ---- robocopy の出力解析パターン ---------------------------
#   サイズ + 日時 + フルパス の行にマッチする。
$LINE_PATTERN = '^\s*(\d+)\s+(\d{4}/\d{2}/\d{2}\s+\d{2}:\d{2}:\d{2})\s+(.*)$'

# ############################################################
#  設定はここまで
# ############################################################


# ------------------------------------------------------------
# スクリプト自身が置かれているフォルダを求める
#
#   カレントディレクトリ (Get-Location) を基準にすると、
#   「どこから実行したか」で動作が変わってしまう。
#   $PSScriptRoot はスクリプト自身の位置を返すため、
#   ツールをどこに置いても、どこから実行しても同じ動作になる。
# ------------------------------------------------------------
$scriptDir = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($scriptDir)) {
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
}
if ([string]::IsNullOrWhiteSpace($scriptDir)) {
    $scriptDir = (Get-Location).ProviderPath
}


# ------------------------------------------------------------
# 走査対象を決める
# ------------------------------------------------------------
if ($null -eq $Path -or $Path.Count -eq 0) {
    if ($DEFAULT_TARGET_PATHS.Count -gt 0) {
        $Path = $DEFAULT_TARGET_PATHS
    } else {
        $Path = @($env:USERPROFILE)
        Write-Host ''
        Write-Host '[お知らせ] 走査対象が指定されていないため、既定値で実行します。' -ForegroundColor Cyan
        Write-Host "          対象: $env:USERPROFILE" -ForegroundColor Cyan
        Write-Host '          フォルダを指定する場合は、スクリプト冒頭の' -ForegroundColor Cyan
        Write-Host '          $DEFAULT_TARGET_PATHS = @( ... ) の中に記述してください。' -ForegroundColor Cyan
        Write-Host '          (行頭に # が付いている行はコメントで、無効です)' -ForegroundColor Cyan
    }
}

# ------------------------------------------------------------
# 対象パスの検証
#
#   共有フォルダは、ネットワークの状態や権限でアクセスできないことがある。
#   1 つでも到達できないと全体が止まる作りにすると、
#   「他は走査できたはずなのに何も出力されない」という最も困る結果になる。
#   到達できないパスはスキップして続行し、その事実をログに必ず残す。
# ------------------------------------------------------------
$validRoots   = @()
$skippedRoots = @()

foreach ($p in $Path) {
    if ([string]::IsNullOrWhiteSpace($p)) { continue }

    if (Test-Path -LiteralPath $p) {
        try {
            $resolved = (Resolve-Path -LiteralPath $p).ProviderPath

            # 末尾の \ を単純に削ると、ドライブ直下 (F:\) が F: になる。
            # robocopy に F: を渡すと「F ドライブのカレントディレクトリ」と
            # 解釈されるおそれがあるため、ドライブ直下は \ を残す。
            if ($resolved -match '^[A-Za-z]:\\$') {
                $scanPath = $resolved
                $basePath = $resolved.TrimEnd('\')
            } else {
                $scanPath = $resolved.TrimEnd('\')
                $basePath = $scanPath
            }

            $validRoots += [PSCustomObject]@{ ScanPath = $scanPath; BasePath = $basePath }

        } catch {
            $skippedRoots += [PSCustomObject]@{ Path = $p; Reason = $_.Exception.Message }
        }
    } else {
        $skippedRoots += [PSCustomObject]@{ Path = $p; Reason = 'パスが存在しないか、アクセスできません' }
    }
}

if ($validRoots.Count -eq 0) {
    Write-Error '走査可能なフォルダが 1 つもありません。設定を確認してください。'
    exit 1
}

# ------------------------------------------------------------
# 出力先を決める
# ------------------------------------------------------------
if ([string]::IsNullOrWhiteSpace($OutputDir)) {
    $OutputDir = if (-not [string]::IsNullOrWhiteSpace($DEFAULT_OUTPUT_DIR)) {
        $DEFAULT_OUTPUT_DIR
    } else {
        $scriptDir
    }
}
if (-not (Test-Path -LiteralPath $OutputDir)) {
    try {
        New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
    } catch {
        Write-Error "出力先フォルダを作成できませんでした: $OutputDir`n$($_.Exception.Message)"
        exit 1
    }
}

$outputDirFull = (Resolve-Path -LiteralPath $OutputDir).ProviderPath.TrimEnd('\')
$now   = Get-Date
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'

# 各区分の境界日時。ここより古い (以前の) ファイルが該当する。
$border1 = $now.AddYears(-1)
$border3 = $now.AddYears(-3)
$border5 = $now.AddYears(-5)

# UTF-8 (BOM 付き)。BOM がないと Excel で開いた際に日本語が文字化けする。
$encoding = New-Object System.Text.UTF8Encoding($true)

$csvPath = @{
    1 = Join-Path $outputDirFull ($CSV_NAME_1Y -f $stamp)
    3 = Join-Path $outputDirFull ($CSV_NAME_3Y -f $stamp)
    5 = Join-Path $outputDirFull ($CSV_NAME_5Y -f $stamp)
}
$logPath = Join-Path $outputDirFull ($LOG_NAME -f $stamp)


# ------------------------------------------------------------
# 実行条件の表示
#
#   区分の境界日を明示する。
#   「2023年のファイルなのに 1年CSV に入っている」といった疑問は、
#   境界が見えないことから生じる。実際には 2026/07 時点で
#   2023/09 のファイルは経過 2年9か月であり、まだ 3年に達していない。
# ------------------------------------------------------------
Write-Host ''
Write-Host '========================================'
Write-Host ' 走査対象 :'
foreach ($r in $validRoots) { Write-Host "   - $($r.ScanPath)" }
if ($skippedRoots.Count -gt 0) {
    Write-Host ''
    Write-Host ' 走査できないためスキップ:' -ForegroundColor Yellow
    foreach ($s in $skippedRoots) {
        Write-Host "   - $($s.Path)  [$($s.Reason)]" -ForegroundColor Yellow
    }
}
Write-Host ''
Write-Host " 出力先   : $outputDirFull"
Write-Host " 並列数   : 最大 $MAX_PARALLEL フォルダ同時"
Write-Host ' 判定基準 : 最終更新日時 (LastWriteTime)'
Write-Host ' 列挙方式 : robocopy /L (リストモード)'
Write-Host ''
Write-Host ' 区分の境界 (重複なし):'
Write-Host ("   1年CSV : {0} より後 〜 {1} 以前" -f $border3.ToString('yyyy-MM-dd'), $border1.ToString('yyyy-MM-dd'))
Write-Host ("   3年CSV : {0} より後 〜 {1} 以前" -f $border5.ToString('yyyy-MM-dd'), $border3.ToString('yyyy-MM-dd'))
Write-Host ("   5年CSV : {0} 以前" -f $border5.ToString('yyyy-MM-dd'))
Write-Host '========================================'
Write-Host ''


# ------------------------------------------------------------
# 走査対象のトップフォルダを組み立てる
#
#   指定された各ルートについて
#     - ルート直下のファイル (LEV:1)
#     - ルート直下の各フォルダ (再帰)
#   を「1つの走査単位」とする。
#   この単位ごとに並列実行し、進捗と集計を出す。
# ------------------------------------------------------------
$targets = @()

foreach ($root in $validRoots) {

    $scanRoot = $root.ScanPath
    $baseRoot = $root.BasePath

    $targets += [PSCustomObject]@{
        Label    = "$scanRoot\(直下)"
        RootKey  = $scanRoot      # ログはこの単位で集計する
        ScanPath = $scanRoot
        BaseRoot = $baseRoot
        TopOnly  = $true
    }

    Get-ChildItem -LiteralPath $scanRoot -Directory -Force -ErrorAction SilentlyContinue |
        Sort-Object Name |
        ForEach-Object {
            $targets += [PSCustomObject]@{
                Label    = $_.FullName
                RootKey  = $scanRoot
                ScanPath = $_.FullName
                BaseRoot = $baseRoot
                TopOnly  = $false
            }
        }
}

Write-Host "走査単位: $($targets.Count) フォルダ"
Write-Host ''


# ============================================================
# 並列ワーカー
#
#   1 つのトップフォルダを robocopy で走査し、
#   該当ファイルを CSV へ書き出して、集計を返す。
#
#   [重要] StreamWriter はスレッドセーフではない。
#   複数のワーカーが同じ CSV へ同時に書き込むため、
#   呼び出し側で TextWriter.Synchronized() によるラッパを渡している。
#   ワーカー側はそれを普通の writer として使えばよい。
# ============================================================
$worker = {
    param(
        $ScanPath, $BaseRoot, $Label, $RootKey, $TopOnly,
        $Writers, $Now, $Border1, $Border3, $Border5, $LinePattern
    )

    # --- CSV フィールドのエスケープ ---
    function ConvertTo-CsvField([string]$Value) {
        if ($null -eq $Value) { $Value = '' }
        return '"' + $Value.Replace('"', '""') + '"'
    }

    # --- 経過年数の区分を決める (5 -> 3 -> 1 の順。最初に該当した1つだけ) ---
    function Get-Bucket([datetime]$LastWrite) {
        if ($LastWrite -le $Border5) { return 5 }
        if ($LastWrite -le $Border3) { return 3 }
        if ($LastWrite -le $Border1) { return 1 }
        return 0
    }

    $stat = @{ Scanned = 0; Hit1 = 0; Hit3 = 0; Hit5 = 0; Bytes = [long]0 }

    # 保留中のレコード (robocopy のパス折り返しに対応するため)
    $pending = @{ Size = $null; Time = $null; Path = $null }

    # --- 1 レコードを判定して書き出す ---
    $emit = {
        param($Size, $TimeText, $FullPath)

        $stat.Scanned++

        $lastWrite = [datetime]::MinValue
        if (-not [datetime]::TryParseExact(
                $TimeText, 'yyyy/MM/dd HH:mm:ss',
                [Globalization.CultureInfo]::InvariantCulture,
                [Globalization.DateTimeStyles]::None,
                [ref] $lastWrite)) {
            return
        }

        $bucket = Get-Bucket $lastWrite
        if ($bucket -eq 0) { return }

        switch ($bucket) {
            1 { $stat.Hit1++ }
            3 { $stat.Hit3++ }
            5 { $stat.Hit5++ }
        }
        $stat.Bytes += $Size

        $name   = Split-Path -Path $FullPath -Leaf
        $folder = Split-Path -Path $FullPath -Parent
        $ext    = [IO.Path]::GetExtension($name)
        if ([string]::IsNullOrEmpty($ext)) { $ext = '(なし)' } else { $ext = $ext.ToLower() }

        $sizeKB      = [math]::Round($Size / 1KB, 1)
        $elapsedDays = [int]($Now - $lastWrite).TotalDays

        $depth = 0
        if ($folder.Length -gt $BaseRoot.Length) {
            $relative = $folder.Substring($BaseRoot.Length).Trim('\')
            if ($relative) { $depth = ($relative -split '\\').Count }
        }

        $fields = @(
            (ConvertTo-CsvField $name)
            (ConvertTo-CsvField $ext)
            (ConvertTo-CsvField ([string]$sizeKB))
            (ConvertTo-CsvField $lastWrite.ToString('yyyy-MM-dd HH:mm'))
            (ConvertTo-CsvField ([string]$elapsedDays))
            (ConvertTo-CsvField ([string]$depth))
            (ConvertTo-CsvField $folder)
        )

        # 同期ラッパ経由なので、複数ワーカーから同時に呼んでも安全
        $Writers[$bucket].WriteLine($fields -join ',')
    }

    # --- robocopy をリストモードで実行 ---
    $roboArgs = @(
        $ScanPath
        'NULL'
        '/L', '/NJH', '/NJS', '/FP', '/NDL', '/NC', '/BYTES', '/TS', '/R:0', '/W:0'
    )
    $roboArgs += if ($TopOnly) { '/LEV:1' } else { '/S' }

    # 出力を 1 行ずつ流す。全体を変数に受けるとメモリを圧迫するため。
    & robocopy.exe @roboArgs 2>$null | ForEach-Object {

        $line = $_
        if ([string]::IsNullOrWhiteSpace($line)) { return }

        if ($line -match $LinePattern) {

            if ($null -ne $pending.Path) {
                & $emit $pending.Size $pending.Time $pending.Path
            }
            $pending.Size = [long]$Matches[1]
            $pending.Time = $Matches[2]
            $pending.Path = $Matches[3].TrimEnd()

        } else {
            # サイズ・日時を持たない行 = 直前のパスの折り返し。連結して復元する。
            if ($null -ne $pending.Path) {
                $pending.Path = $pending.Path + $line.Trim()
            }
        }
    }

    if ($null -ne $pending.Path) {
        & $emit $pending.Size $pending.Time $pending.Path
    }

    # 集計を返す
    [PSCustomObject]@{
        Label   = $Label
        RootKey = $RootKey
        Scanned = $stat.Scanned
        Hit1    = $stat.Hit1
        Hit3    = $stat.Hit3
        Hit5    = $stat.Hit5
        Bytes   = $stat.Bytes
    }
}


# ============================================================
# 出力ストリームを開き、並列で走査する
# ============================================================

$rawWriters = @{}   # 実体 (最後に閉じる)
$writers    = @{}   # 同期ラッパ (ワーカーへ渡す)
$logWriter  = $null
$pool       = $null

try {
    foreach ($y in 1, 3, 5) {
        $raw = New-Object System.IO.StreamWriter($csvPath[$y], $false, $encoding, $BUFFER_SIZE)
        $raw.AutoFlush = $false
        $raw.WriteLine($CSV_HEADER)

        $rawWriters[$y] = $raw

        # StreamWriter はスレッドセーフではない。
        # 複数のワーカーが同じ CSV へ同時に書き込むため、同期ラッパで包む。
        # これを怠ると、行が混ざる・欠落する・例外になる。
        $writers[$y] = [System.IO.TextWriter]::Synchronized($raw)
    }

    $logWriter = New-Object System.IO.StreamWriter($logPath, $false, $encoding, $BUFFER_SIZE)
    $logWriter.AutoFlush = $false

    $logWriter.WriteLine("走査開始 : $($now.ToString('yyyy-MM-dd HH:mm:ss'))")
    $logWriter.WriteLine("並列数   : 最大 $MAX_PARALLEL")
    $logWriter.WriteLine('走査対象 :')
    foreach ($r in $validRoots) { $logWriter.WriteLine("  - $($r.ScanPath)") }
    if ($skippedRoots.Count -gt 0) {
        $logWriter.WriteLine('')
        $logWriter.WriteLine('スキップした対象 (走査できませんでした) :')
        foreach ($s in $skippedRoots) {
            $logWriter.WriteLine("  - $($s.Path)  [$($s.Reason)]")
        }
    }
    $logWriter.WriteLine('')
    $logWriter.WriteLine('区分の境界 (重複なし) :')
    $logWriter.WriteLine(("  1年CSV : {0} より後 〜 {1} 以前" -f $border3.ToString('yyyy-MM-dd'), $border1.ToString('yyyy-MM-dd')))
    $logWriter.WriteLine(("  3年CSV : {0} より後 〜 {1} 以前" -f $border5.ToString('yyyy-MM-dd'), $border3.ToString('yyyy-MM-dd')))
    $logWriter.WriteLine(("  5年CSV : {0} 以前" -f $border5.ToString('yyyy-MM-dd')))
    $logWriter.WriteLine('')
    $logWriter.WriteLine('指定フォルダ別 集計')
    $logWriter.WriteLine('--------------------------------------------------------------------------')

    # --------------------------------------------------------
    # RunspacePool を作る
    #
    #   Start-Job は 1 ジョブごとに PowerShell のプロセスを起動するため、
    #   数十フォルダを回すには起動コストが大きすぎる。
    #   RunspacePool は同一プロセス内でスレッドを使い回すため軽量。
    #
    #   上限を $MAX_PARALLEL に絞る。
    #   ディスクの読み取り速度が上限に達すると、並列数を増やしても
    #   待ち時間が増えるだけで速くならない。
    # --------------------------------------------------------
    $pool = [runspacefactory]::CreateRunspacePool(1, $MAX_PARALLEL)
    $pool.Open()

    $jobs = @()

    foreach ($t in $targets) {
        $ps = [powershell]::Create()
        $ps.RunspacePool = $pool

        [void]$ps.AddScript($worker).
            AddArgument($t.ScanPath).
            AddArgument($t.BaseRoot).
            AddArgument($t.Label).
            AddArgument($t.RootKey).
            AddArgument($t.TopOnly).
            AddArgument($writers).
            AddArgument($now).
            AddArgument($border1).
            AddArgument($border3).
            AddArgument($border5).
            AddArgument($LINE_PATTERN)

        $jobs += [PSCustomObject]@{
            Label  = $t.Label
            Shell  = $ps
            Handle = $ps.BeginInvoke()
        }
    }

    # --------------------------------------------------------
    # 完了を待ちながら進捗を表示する
    #
    #   画面には走査したフォルダごとに進捗を出すが、
    #   ログに書くのは「指定したフォルダ (ルート) 単位の集計」のみ。
    #   配下の内訳は CSV を見れば分かるため、ログには残さない。
    #   ログはあくまで「どの対象を、どれだけ走査したか」の記録に徹する。
    # --------------------------------------------------------
    $grand     = @{ Scanned = 0; Hit1 = 0; Hit3 = 0; Hit5 = 0; Bytes = [long]0 }
    $done      = 0
    $total     = $jobs.Count
    $collected = @()

    # ルート単位の集計。キーは指定されたフォルダのパス。
    $rootStats = @{}
    foreach ($r in $validRoots) {
        $rootStats[$r.ScanPath] = @{ Scanned = 0; Hit1 = 0; Hit3 = 0; Hit5 = 0; Bytes = [long]0 }
    }

    while ($collected.Count -lt $total) {

        foreach ($job in $jobs) {

            if ($job.Handle.IsCompleted -and ($collected -notcontains $job.Label)) {

                try {
                    $result = $job.Shell.EndInvoke($job.Handle)
                } catch {
                    Write-Host ("  走査失敗 : {0}  [{1}]" -f $job.Label, $_.Exception.Message) -ForegroundColor Yellow
                    $result = $null
                }

                $job.Shell.Dispose()
                $collected += $job.Label
                $done++

                if ($null -ne $result -and $result.Count -gt 0) {
                    $r = $result[0]

                    # 画面には進捗として表示する
                    Write-Host ("[{0,3}/{1}] {2}" -f $done, $total, $r.Label)

                    # ルート単位へ加算する (ログはこの単位でのみ出力する)
                    $rs = $rootStats[$r.RootKey]
                    $rs.Scanned += $r.Scanned
                    $rs.Hit1    += $r.Hit1
                    $rs.Hit3    += $r.Hit3
                    $rs.Hit5    += $r.Hit5
                    $rs.Bytes   += $r.Bytes

                    $grand.Scanned += $r.Scanned
                    $grand.Hit1    += $r.Hit1
                    $grand.Hit3    += $r.Hit3
                    $grand.Hit5    += $r.Hit5
                    $grand.Bytes   += $r.Bytes
                }
            }
        }

        if ($collected.Count -lt $total) { Start-Sleep -Milliseconds 200 }
    }

    # --------------------------------------------------------
    # 指定フォルダ (ルート) 単位の集計をログへ書き出す
    # --------------------------------------------------------
    foreach ($r in $validRoots) {
        $rs = $rootStats[$r.ScanPath]
        $mb = [math]::Round($rs.Bytes / 1MB, 1)

        $logWriter.WriteLine($r.ScanPath)
        $logWriter.WriteLine(('    走査ファイル数 : {0,10} 件' -f $rs.Scanned))
        $logWriter.WriteLine(('    1年以上3年未満 : {0,10} 件' -f $rs.Hit1))
        $logWriter.WriteLine(('    3年以上5年未満 : {0,10} 件' -f $rs.Hit3))
        $logWriter.WriteLine(('    5年以上        : {0,10} 件' -f $rs.Hit5))
        $logWriter.WriteLine(('    該当合計サイズ : {0,10} MB' -f $mb))
        $logWriter.WriteLine('')
    }

    $totalMB  = [math]::Round($grand.Bytes / 1MB, 1)
    $totalHit = $grand.Hit1 + $grand.Hit3 + $grand.Hit5

    $logWriter.WriteLine('--------------------------------------------------------------------------')
    $logWriter.WriteLine("走査ファイル数 : $($grand.Scanned) 件")
    $logWriter.WriteLine("該当ファイル数 : $totalHit 件")
    $logWriter.WriteLine("  1年以上3年未満 : $($grand.Hit1) 件")
    $logWriter.WriteLine("  3年以上5年未満 : $($grand.Hit3) 件")
    $logWriter.WriteLine("  5年以上        : $($grand.Hit5) 件")
    $logWriter.WriteLine("該当合計サイズ : $totalMB MB")
    if ($skippedRoots.Count -gt 0) {
        $logWriter.WriteLine("スキップした対象 : $($skippedRoots.Count) 件 (上記参照)")
    }
    $logWriter.WriteLine("走査終了 : $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))")

    Write-Host ''
    Write-Host '========================================'
    Write-Host " 走査ファイル数 : $($grand.Scanned) 件"
    Write-Host " 該当ファイル数 : $totalHit 件"
    Write-Host "   1年以上3年未満 : $($grand.Hit1) 件"
    Write-Host "   3年以上5年未満 : $($grand.Hit3) 件"
    Write-Host "   5年以上        : $($grand.Hit5) 件"
    Write-Host " 該当合計サイズ : $totalMB MB"
    if ($skippedRoots.Count -gt 0) {
        Write-Host " スキップ       : $($skippedRoots.Count) 件（走査できませんでした）" -ForegroundColor Yellow
    }
    Write-Host '========================================'
}
finally {
    # --------------------------------------------------------
    # 後始末
    #
    #   バッファリングしているため、閉じ忘れるとバッファに残った
    #   データが失われる。例外が発生した場合も必ず書き出す。
    #
    #   同期ラッパではなく、実体 (rawWriters) を閉じる。
    # --------------------------------------------------------
    if ($null -ne $pool) {
        $pool.Close()
        $pool.Dispose()
    }
    foreach ($w in $rawWriters.Values) {
        if ($null -ne $w) { $w.Flush(); $w.Close(); $w.Dispose() }
    }
    if ($null -ne $logWriter) {
        $logWriter.Flush(); $logWriter.Close(); $logWriter.Dispose()
    }
}

Write-Host ''
Write-Host '出力しました:'
foreach ($y in 1, 3, 5) { Write-Host "  $($csvPath[$y])" }
Write-Host "  $logPath"
Write-Host ''
Write-Host '※ 本ツールはファイルの削除・移動を行いません。'
Write-Host '  CSV を確認のうえ、削除の可否はご自身でご判断ください。'
Write-Host ''
