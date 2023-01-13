#include ..\Lib\Acc.ahk
#Requires AutoHotkey v2.0-a

oCaretPos := {x:0,y:0}
Loop {
    try oCaretPos := Acc.ElementFromHandle("A", Acc.ObjId.Caret).Location
    if oCaretPos.x=0 && oCaretPos.y=0
        CaretGetPos(&x, &y), ToolTip("x: " x " y: " y)
    else
        ToolTip("x: " oCaretPos.x " y: " oCaretPos.y)
}