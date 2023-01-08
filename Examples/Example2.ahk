#include ..\Lib\Acc.ahk
#Requires AutoHotkey v2.0-a

Run("notepad.exe")
WinWaitActive("ahk_exe notepad.exe")
oAcc := Acc.ObjectFromWindow("ahk_exe notepad.exe")
oEditor := oAcc[4,1,4] ; Can get the element by path

; If the editor element wasn't found, then we are probably dealing with Windows 11
if oEditor.RoleText != "editable text" { 
    oEditor := oAcc.FindFirst({RoleText:"editable text"})
    Win11 := True
}
oEditor.Highlight()
oEditor.Value := "Example text" ; Set the value

oAcc.FindFirst({Name:"File"}).DoDefaultAction()
; Windows 11 menu is in another window...
if IsSet(Win11) {
    WinWait("PopupHost ahk_exe notepad.exe")
    oAcc := Acc.ObjectFromWindow()
}
oAcc.WaitElementExist({Name:"save as", casesensitive:False, matchmode:1}).DoDefaultAction() ; Wait "Save As" to exist, matching the start of string, case insensitive