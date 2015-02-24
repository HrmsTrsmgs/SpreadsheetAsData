@echo off
if exist "%systemroot%\sysnative\" (
ansicon\x86\ansicon
) else if exist "%systemroot%\syswow64\" (
ansicon\x64\ansicon
) else (
ansicon\x86\ansicon
)

mode con:cols=140
