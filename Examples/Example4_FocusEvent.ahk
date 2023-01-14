﻿#include ..\Lib\Acc.ahk
#Requires AutoHotkey v2.0-a

Persistent()

; This won't work without assigning the registered event object to a variable
; because the event will be deregistered once the object is destroyed (and not assigning it to
; any thing will cause it to be destroyed automatically)
h := Acc.RegisterWinEvent(OnFocus, Acc.Event.Object_Focus)
h2 := Acc.RegisterWinEvent(OnMoveSizeStart, Acc.Event.System_MoveSizeStart)

OnFocus(oAcc, info) {
    try ToolTip("Element: " oAcc.Name "`nEvent: " Acc.Event[info.Event] "`nTime: " info.EventTime 
        . "`nSender WinID: " info.WinID "`nSender ControlID: " info.ControlID)
}
OnMoveSizeStart(oAcc, info) {
    try ToolTip("Element: " oAcc.Name "`nEvent: " Acc.Event[info.Event])
}
