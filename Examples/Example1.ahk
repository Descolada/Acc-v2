#include ..\Lib\Acc.ahk

Run("notepad.exe")
WinWaitActive("ahk_exe notepad.exe")
oAcc := Acc.ObjectFromWindow("ahk_exe notepad.exe")
A_Clipboard := oAcc.DumpAll()
MsgBox("Notepad element dumped into clipboard!")