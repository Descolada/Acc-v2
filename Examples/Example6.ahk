#include ..\Lib\Acc.ahk
#Requires AutoHotkey v2.0-a

SetTitleMatchMode(2)

Run("explore C:\")
CDriveName := DriveGetLabel("C:") " (C:)"
WinWaitActive(CDriveName)

oExplorer := Acc.ObjectFromWindow()
oExplorer.FindFirst({Name:"Program Files", RoleText:"list item", matchmode:2}).Select(Acc.SELECTIONFLAG.ADDSELECTION)
(oWin := oExplorer.FindFirst({Name:"Windows", RoleText:"list item"})).Select(Acc.SELECTIONFLAG.ADDSELECTION)

selectedItems := ""
for i, selected in oWin.Parent.Selection
    selectedItems .= "Selection " i ": " selected.Dump() "`n"

MsgBox(selectedItems)
