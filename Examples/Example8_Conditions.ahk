#include ..\Lib\Acc.ahk
#Requires AutoHotkey v2.0-a

Run("notepad.exe")
WinWaitActive("ahk_exe notepad.exe")
oAcc := Acc.ElementFromHandle("ahk_exe notepad.exe")

; And condition is created with {}
MsgBox("First element matching Name `"Maximize`" and Role `"12`":`n" 
    . oAcc.FindElement({Name:"Maximize", Role:12}).Dump())

; Or condition is created with []
result := "All elements matching either Name `"Minimize`" or Name `"Maximize`":`n`n"
for i, v in oAcc.FindElements([{Name:"Minimize"}, {Name:"Maximize"}])
    result .= i ": " v.Dump() "`n"
MsgBox result

; Negate with "not"
result := "All elements matching Name `"Minimize`":`n`n"
for i, v in oAcc.FindElements({Name:"Minimize"})
    result .= i ": " v.Dump() "`n"
result .= "`nAll elements matching Name `"Minimize`" but not Role `"43`":`n`n"
for i, v in oAcc.FindElements({Name:"Minimize", not:{Role:43}})
    result .= i ": " v.Dump() "`n"
MsgBox result

; Specify case-sensitivity with "casesensitive" or "cs", and matchmode with "matchmode" or "mm".
; Case-sensitivity is ignored with numeric values.
result := "All elements matching Name `"minim`" (case insensitive, anywhere in string):`n`n"
for i, v in oAcc.FindElements({Name:"minim", cs:false, mm:2}) ; mm:2 is equal to mm:"SubString"
    result .= i ": " v.Dump() "`n"
MsgBox result

; Find nth element by specifying key "index" or "i". Negative index reverses the search direction.
result := "All elements matching Role `"9`":`n`n"
for i, v in oAcc.FindElements({Role:9},5)
    result .= i ": " v.Dump() "`n"
result .= "`nLast element matching Role `"9`":`n" 
result .= oAcc.FindElement({Role:9, index:-1}).Dump()
result .= "`nSecond element matching Role `"9`":`n" 
result .= oAcc.FindElement({Role:9, index:2}).Dump()
MsgBox result

; To use "or" condition with other options (such as "index"), put it inside an object with key "or"
MsgBox("Last element matching Name `"Save	Ctrl+S`" or Name `"Save As...	Ctrl+Shift+S`":`n" 
    . oAcc.FindElement({or:[{Name:"Save	Ctrl+S"}, {Name:"Save As...	Ctrl+Shift+S"}], index:-1}).Dump())

; To find specific element object (for example to determine the path) use IsEqual
oEdit := oAcc[4,1,4]
MsgBox("oEdit element: " oEdit.Dump()
    . "`n`nFirst element matching oEdit: " oAcc.FindElement({IsEqual:oEdit}).Dump()
    . "`n`nFound elements path: " oAcc.FindElement({IsEqual:oEdit}).Path)

; Conditions can be used inside paths as well.
MsgBox("oEdit element (4th element with State 1048576, then 1st element, then 1st element with Name `"Text Editor`"): `n`n" 
    . oAcc[{State:1048576, i:4},1,{Name:"Text Editor"}].Dump())