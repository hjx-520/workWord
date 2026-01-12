@echo off
setlocal enableDelayedExpansion

REM Date for ABCD filename
REM Task scheduler set FileDate=%date:~12,2%%date:~7,2%%date:~4,2%
set FileDate=%date:~12,2%%date:~4,2%%date:~7,2%

REM Current month
REM Task scheduler set m=%date:~7,2%
set m=%date:~4,2%

if %m%==01 set monthName=Jan
if %m%==02 set monthName=Feb
if %m%==03 set monthName=Mar
if %m%==04 set monthName=Apr
if %m%==05 set monthName=May
if %m%==06 set monthName=Jun
if %m%==07 set monthName=Jul
if %m%==08 set monthName=Aug
if %m%==09 set monthName=Sep
if %m%==10 set monthName=Oct
if %m%==11 set monthName=Nov
if %m%==12 set monthName=Dec

REM Month name of backup folder
REM Task scheduler set SrcDate=%date:~10,4%%monthName%%date:~4,2%
set SrcDate=%date:~10,4%%monthName%%date:~7,2%

set src_file=D:\TIHostInter\FTP\BACKUP\%SrcDate%\INBOUND\BEFORE\
REM PRD set src_file=D:\TIHostInter\FTP\BACKUP\%SrcDate%\INBOUND\

copy %src_file%BILTOTFS.BTF05BXR.D%FileDate%.TXT "D:\backup\ABCDFile\BILTOTFS.BTF05BXR.D%FileDate%.TXT"
