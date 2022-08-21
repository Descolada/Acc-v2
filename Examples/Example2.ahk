#include ..\Lib\Acc.ahk

Run("notepad.exe")
WinActivate("ahk_exe notepad.exe")
WinWaitActive("ahk_exe notepad.exe")
oAcc := Acc.ObjectFromWindow("ahk_exe notepad.exe")
oEditor := oAcc[4,1,4] ; Can get the element by path
oEditor.Highlight(2000)
oEditor.Value := "Example text" ; Set the value

oAcc.FindFirst({Name:"File"}).DoDefaultAction()
oAcc.WaitElementExist({Name:"save as...", casesensitive:False, matchmode:1}).DoDefaultAction() ; Wait "Save As..." to exist, matching the start of string, case insensitive