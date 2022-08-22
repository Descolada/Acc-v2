#include ..\Lib\Acc.ahk

lnCol := GetVSCodeLnCol()
MsgBox("Ln " lnCol.Ln ", Col " lnCol.Col "`nFound elements path: " lnCol.Path)

GetVSCodeLnCol() {
    oLnCol := Acc.ObjectFromWindow("ahk_exe Code.exe").FindFirst({Name:"Ln \d+, Col \d+", matchmode:"Regex"})
    RegExMatch(oLnCol.Name, "Ln (\d+), Col (\d+)", &match)
    return {Ln:match[1], Col:match[2], Path:oLnCol.Path}
}