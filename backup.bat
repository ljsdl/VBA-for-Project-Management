echo on
echo �ۑ���̃t�H���_�p�X���J��
CD /D X:\

MD Backup
CD X:\Backup

echo �i�[�p�̓��t�t�H���_���쐬����
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

echo �i�[�p�̃X�R�[�v�Ǘ��t�H���_���쐬����
MD 1_�X�R�[�v�Ǘ�
echo �i�[�p�̐i���Ǘ��t�H���_���쐬����
MD 2_�i���Ǘ�

echo off
echo 1_�X�R�[�v�Ǘ�-------------------------------START

echo ************************************************
XCOPY /y "Z:\VBA\TOOLS\*.xls*" 1_�X�R�[�v�Ǘ�
echo ************************************************
echo 1_�X�R�[�v�Ǘ�--------------------------------END

echo 2_�i���Ǘ�-----------------------------------START
echo ************************************************
XCOPY /y "Z:\VBA\�ȑO��\Excel vba��?�n��?��.doc" 2_�i���Ǘ�
echo ************************************************
echo 2_�i���Ǘ�-----------------------------------END

echo �o�b�N�A�b�v�������܂����B
set OPEN_DIR=X:\Backup\%filename%
EXPLORER %OPEN_DIR%
pause


