#include ..\Lib\Acc.ahk
oCaretPos := {x:0,y:0}
Loop {
    try oCaretPos := Acc.ObjectFromWindow(, Acc.OBJID.CARET).Location
    if oCaretPos.x=0 && oCaretPos.y=0
        CaretGetPos(&x, &y), ToolTip("x: " x " y: " y)
    else
        ToolTip("x: " oCaretPos.x " y: " oCaretPos.y)
}