#include ..\Lib\Acc.ahk

Run("notepad.exe")
WinWaitActive("ahk_exe notepad.exe")
oAcc := Acc.ObjectFromWindow("ahk_exe notepad.exe")
st := A_TickCount
Loop 10000 {
oAcc.ValidateCondition({Name:"Test", matchmode:"Regex"})
}
OutputDebug(A_TickCount-st)