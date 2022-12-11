#include ..\Lib\Acc.ahk
#Requires AutoHotkey v2.0-a

Persistent()

global handler

MsgBox("Press F1 to start capturing all system events.`nPress F2 to stop capturing.`nPress Esc to exit app.")

/*
    This function will be called by Acc when a system event happens. 
    Multiple events can be registered to one callback function.
    If we were interested only in getting oAcc, we could use OnSystemEvent(oAcc, *) instead.
*/
OnSystemEvent(oAcc, event, time) {
    try {
        ToolTip("Element: " oAcc.Name "`nEvent: " Acc.EVENT[event] "`nTime: " time)
        SetTimer(ToolTip, 2000)
    }
}
/*
    Check whether the system events have already been registered to the handler variable.
    If not, then register an event range from EVENT.SYSTEM_SOUND to EVENT.SYSTEM_MINIMIZEEND,
        which will also include all the event inbetween those two (eg SYSTEM_CAPTURESTART).

    Note that the events will be deregistered once the returned handler object is destroyed,
        so we always need to store it in a variable, in this case the "handler" variable.
        Also it can't be a local variable, because local variables get destroyed after
        exiting the function/hotkey.
*/
F1::
{
    global handler
    if !IsSet(handler)
        handler := Acc.RegisterWinEvent(Acc.EVENT.SYSTEM_SOUND, Acc.Event.SYSTEM_MINIMIZEEND, OnSystemEvent)
}
; Deregistering the events will happen when the event handler object is destroyed, for example by unsetting it
F2::global handler := unset
Esc::ExitApp