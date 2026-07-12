@echo off
rem ============================================================
rem  Find-StaleFiles 起動用バッチ
rem
rem  PowerShell スクリプトは、既定の実行ポリシー (Restricted) では
rem  実行できない。このバッチは、システムの設定を変更せず、
rem  この 1 回の実行に対してだけポリシーを回避して起動する。
rem
rem   -ExecutionPolicy Bypass : この実行だけポリシーを無視する
rem                             (システムの設定は変更しない)
rem   -NoProfile              : プロファイルを読み込まない (起動が速く、環境の影響を受けない)
rem   -File                   : スクリプトを指定して実行する
rem
rem   %~dp0 : このバッチが置かれているフォルダ
rem           カレントディレクトリではないため、どこから実行しても
rem           同じフォルダのスクリプトを確実に呼び出せる。
rem
rem   %*    : このバッチに渡された引数をそのまま PowerShell へ渡す
rem ============================================================

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Find-StaleFiles.ps1" %*

echo.
echo 終了するには何かキーを押してください...
pause > nul
