#include ..\Lib\Acc.ahk
#Requires AutoHotkey v2.0-a

SetTitleMatchMode(2)

Run("explore C:\")
CDriveName := DriveGetLabel("C:") " (C:)"
WinWaitActive(CDriveName)

oExplorer := Acc.ElementFromHandle()
oExplorer.FindElement({Name:"Program Files", RoleText:"list item", matchmode:"Substring"}).Select("AddSelection")
(oWin := oExplorer.FindElement({Name:"Windows", RoleText:"list item"})).Select(Acc.SelectionFlag.AddSelection)

selectedItems := ""
for i, selected in oWin.Parent.Selection
    selectedItems .= "Selection " i ": " selected.Dump() "`n"

MsgBox(selectedItems)
