#include ..\Lib\Acc.ahk
#Requires AutoHotkey v2.0-a

Persistent()

; This won't work without assigning the registered event object to a variable
; because the event will be deregistered once the object is destroyed (and not assigning it to
; any thing will cause it to be destroyed automatically)
h := Acc.RegisterWinEvent(Acc.EVENT.OBJECT_FOCUS, OnFocus)
h2 := Acc.RegisterWinEvent(Acc.EVENT.SYSTEM_MOVESIZESTART, OnMoveSizeStart)

OnFocus(oAcc, info) {
    try ToolTip("Element: " oAcc.Name "`nEvent: " Acc.EVENT[info.Event] "`nTime: " info.EventTime 
        . "`nSender WinID: " info.WinID "`nSender ControlID: " info.ControlID)
}
OnMoveSizeStart(oAcc, info) {
    try ToolTip("Element: " oAcc.Name "`nEvent: " Acc.EVENT[info.Event])
}
