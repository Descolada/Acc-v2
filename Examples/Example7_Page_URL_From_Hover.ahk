#include ..\Lib\Acc.ahk
#Requires AutoHotkey v2.0-a

; Returns the URL from the document element of a web browser. 
; Hover over a browser with the cursor to test it out.

Loop {
    oAcc := Acc.ElementFromPoint(), oEl := ""
    try {
        oEl := oAcc.Normalize({RoleText:"document", not:{Value:""}})
        RoleText := "`"`"", Name := " `"`""
        try RoleText := "`"" oAcc.RoleText "`""
        try Name := " `"" oAcc.Name "`""
        ToolTip("Document element value: " oEl.Value "`nActual element under mouse: " RoleText Name)
    }
}