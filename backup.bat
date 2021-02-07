echo on
echo 保存先のフォルダパスを開く
CD /D X:\

MD Backup
CD X:\Backup

echo 格納用の日付フォルダを作成する
echo %date%
set yyyy=%date:~0,4%
set mm=%date:~5,2%
set dd=%date:~8,2%
echo %yyyy%
echo %mm%
echo %dd%

echo %time%
set time2=%time: =0%

echo %time2%

set hh=%time2:~0,2%
set mn=%time2:~3,2%
set ss=%time2:~6,2%
echo %hh%
echo %mn%
echo %ss%

set filename=%yyyy%%mm%%dd%_%hh%%mn%%ss%
MD %filename%
CD %filename%

echo 格納用のスコープ管理フォルダを作成する
MD 1_スコープ管理
echo 格納用の進捗管理フォルダを作成する
MD 2_進捗管理

echo off
echo 1_スコープ管理-------------------------------START

echo ************************************************
XCOPY /y "Z:\VBA\TOOLS\*.xls*" 1_スコープ管理
echo ************************************************
echo 1_スコープ管理--------------------------------END

echo 2_進捗管理-----------------------------------START
echo ************************************************
XCOPY /y "Z:\VBA\以前の\Excel vba入?系列?座.doc" 2_進捗管理
echo ************************************************
echo 2_進捗管理-----------------------------------END

echo バックアップ成功しました。
set OPEN_DIR=X:\Backup\%filename%
EXPLORER %OPEN_DIR%
pause


