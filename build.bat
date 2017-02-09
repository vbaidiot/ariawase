@echo off
cd %~dp0
c2def.exe
cscript //nologo vbac.wsf combine
