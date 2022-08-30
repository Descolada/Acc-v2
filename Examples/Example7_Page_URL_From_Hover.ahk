#include ..\Lib\Acc.ahk

; Returns the URL from the document element of a web browser. 
; Hover over a browser with the cursor to test it out.

Loop {
    oAcc := Acc.ObjectFromPoint(), oEl := ""
    try {
        oEl := oAcc.Normalize({RoleText:"document", not:{Value:""}})
        RoleText := "`"`"", Name := " `"`""
        try RoleText := "`"" oAcc.RoleText "`""
        try Name := " `"" oAcc.Name "`""
        ToolTip("Document element value: " oEl.Value "`nActual element under mouse: " RoleText Name)
    }
}