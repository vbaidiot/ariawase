@echo off
cd %~dp0
c2sjis.exe
cscript //nologo vbac.wsf combine
