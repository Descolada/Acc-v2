#include ..\Lib\Acc.ahk
#Requires AutoHotkey v2.0-a

Run("notepad.exe")
WinWaitActive("ahk_exe notepad.exe")
oAcc := Acc.ElementFromHandle("ahk_exe notepad.exe")
A_Clipboard := oAcc.DumpAll()
MsgBox("Notepad element dumped into clipboard!")