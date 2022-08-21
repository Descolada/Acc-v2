/*
    Accessibility library for AHK v2

    Authors: Sean, jethrow, Sancarn (v1 code), Descolada
    Special thanks to Lexikos for many tips about v2

    Short introduction:
        Acc v2 in not a port of v1, but instead a complete redesign to incorporate more object-oriented approaches. 

        Notable changes:
        1) All Acc elements are now array-like objects, where the "Length" property contains the number of children, any nth children can be accessed with element[n], and children can be iterated over with for loops.
        2) Acc main functions are contained in the global Acc object
        3) Element methods are contained inside element objects
        4) Element properties can be used without the "acc" prefix
        5) ChildIds have been removed (are handled in the backend), but can be accessed through 
        el.ChildId
        6) Additional methods have been added for elements, such as FindFirst, FindAll, Click
        7) Acc constants are included in the Acc object
        8) AccViewer is built into the library: when ran directly the AccViewer will show, when included
        in another script then it won't show (but can be accessed by calling Acc.Viewer())

    Acc constants/properties:
        Constants can be accessed as properties (eg Acc.OBJID.CARET), or the property name can be
        accessed by getting as an item (eg Acc.OBJID[0xFFFFFFF8])

        OBJID
        STATE
        ROLE
        NAVDIR
        SELECTIONFLAG
        EVENT
    
    Acc methods:
        ObjectFromPoint(x:=unset, y:=unset, &idChild := "")
        ObjectFromWindow(hWnd:="A", idObject := 0)
        GetRootElement()
        RegisterWinEvent(event, callback) 
            Registers an event from Acc.EVENT to a callback function and returns a new object
                containing the WinEventHook
            The callback function needs to have three arguments: 
                CallbackFunction(oAcc, Event, EventTime)
            Unhooking of the event handler will happen once the returned object is destroyed
            (either when overwritten by a constant, or when the script closes).
        SetWinEventHook(eventMin, eventMax, pCallback)
        UnhookWinEvent(hHook)

        Legacy methods:
        ObjectFromPath(ChildPath, hWnd:="A")    => same as ObjectFromWindow[comma-separated path]
        GetRoleText(nRole)                      => same as element.RoleText
        GetStateText(nState)                    => same as element.StateText
        Query(pAcc)                             => for internal use

    IAccessible element properties:
        Name                => Gets or sets the name
        Value               => Gets or sets the value
        Role                => Role as an integer
        RoleText            => Role converted into text form
        Help
        KeyboardShortcut
        State               => State as an integer
        StateText           => State converted into text form
        Description
        DefaultAction
        Focus
        Selection
        Parent              => Returns the parent element
        IsChild             => Checks whether the current element is of child type
        Length              => Returns the number of children the element has
        Location            => Returns an object containing {x,y,w,h}
        Children            => Returns all children as an array (usually not required)
        wId                 => ID of the window the element is inside
        oAcc                => ComObject of the underlying IAccessible
        childId             => childId of the underlying IAccessible
    
    IAccessible element methods:
        Select(flags)       => flags can be any of the SELECTIONFLAG constants
        DoDefaultAction()
        HitTest(x, y)       => unused
        GetNthChild(n)      => Equal to element[n]
        GetLocation(relativeTo:="")
            Returns an object containing the x, y coordinates and width and height: {x:x coordinate, y:y coordinate, w:width, h:height}. relativeTo can be client, window or screen, default is A_CoordModeMouse.
        IsEqual(oCompare)   => Checks whether the element is equal to another element (oCompare)
        FindFirst(condition, scope:=4) 
            Finds the first element matching the condition (see description under ValidateCondition)
            Scope is the search scope: 1=element itself; 2=direct children; 4=descendants (including children of children)
        FindAll(condition, scope:=4)
            Returns an array of elements matching the condition (see description under ValidateCondition)
        WaitElementExist(conditionOrPath, scope:=4, timeOut:=-1)
            Waits an element exist that matches a condition or a path. 
            Timeout less than 1 waits indefinitely, otherwise is the wait time in milliseconds
            A timeout throws an error, otherwise the matching element is returned.
        ValidateCondition(condition)
            Checks whether the element matches a provided condition.
            Everything inside {} is an "and" condition
            Everything inside [] is an "or" condition
            Object key "not" creates a not condition

            matchmode key defines the MatchMode: 1=must start with; 2=can contain anywhere in string; 3=exact match; RegEx

            casesensitive key defines case sensitivity: True=case sensitive; False=case insensitive

            {Name:"Something"} => Name must match "Something" (case sensitive)
            {Name:"Something", matchmode:2, casesensitive:False} => Name must contain "Something" anywhere inside the Name, case insensitive
            {Name:"Something", RoleText:"something else"} => Name must match "Something" and RoleText must match "something else"
            [{Name:"Something", Role:42}, {Name:"Something2", RoleText:"something else"}] => Name=="Something" and Role==42 OR Name=="Something2" and RoleText=="something else"
            {Name:"Something", not:[RoleText:"something", RoleText:"something else"]} => Name must match "something" and RoleText cannot match "something" nor "something else"
        Dump()      => Outputs relevant information about the element (Name, Value, Location etc)
        DumpAll()   => Outputs relevant information about the element and all descendants of the element
        Highlight(showTime:=unset, color:="Red", d:=2)
            Highlights the element for a chosen period of time
            Possible showTime values:
                Unset: removes the highlighting
                0: Indefinite highlighting
                Positive integer (eg 2000): will highlight and pause for the specified amount of time in ms
                Negative integer: will highlight for the specified amount of time in ms, but script execution will continue
            color can be any of the Color names or RGB values
            d sets the border width
        Click(WhichButton:="left", ClickCount:=1, DownOrUp:="", Relative:="")
            Click the center of the element.
            If WhichButton is a number, then Sleep will be called with that number. Eg Click(200) will sleep 200ms after clicking
            If ClickCount is a number >=10, then Sleep will be called with that number. To click 10+ times and sleep after, specify "ClickCount SleepTime". Ex: Click("left", 200) will sleep 200ms after clicking. Ex: Click("left", "20 200") will left-click 20 times and then sleep 200ms.
            If Relative is "Rel" or "Relative" then X and Y coordinates are treated as offsets from the current mouse position. Otherwise it expects offset values for both X and Y (eg "-5 10" would offset X by -5 and Y by +10).
        ControlClick(WhichButton:="left", ClickCount:=1, Options:="")
            ControlClicks the element after getting relative coordinates with GetLocation("client"). 
            If WhichButton is a number, then a Sleep will be called afterwards. Ex: ControlClick(200) will sleep 200ms after clicking. Same for ControlClick("ahk_id 12345", 200)


*/

#DllLoad oleacc

if (!A_IsCompiled and A_LineFile=A_ScriptFullPath)
    Acc.Viewer()

class Acc {
    static PropertyFromValue(obj, value) {
        for k, v in obj.OwnProps()
            if value == v
                return k
        throw UnsetItemError("Property item `"" value "`" not found!", -1)
    }
    static PropertyValueGetter := {get: (obj, value) => Acc.PropertyFromValue(obj, value)}
    static RegisteredWinEvents := Map()

    ;https://msdn.microsoft.com/en-us/library/windows/desktop/dd373606(v=vs.85).aspx
    static OBJID := {WINDOW:0x00000000
        , SYSMENU:0xFFFFFFFF
        , TITLEBAR:0xFFFFFFFE
        , MENU:0xFFFFFFFD
        , CLIENT:0xFFFFFFFC
        , VSCROLL:0xFFFFFFFB
        , HSCROLL:0xFFFFFFFA
        , SIZEGRIP:0xFFFFFFF9
        , CARET:0xFFFFFFF8
        , CURSOR:0xFFFFFFF7
        , ALERT:0xFFFFFFF6
        , SOUND:0xFFFFFFF5
        , QUERYCLASSNAMEIDX:0xFFFFFFF4
        , NATIVEOM:0xFFFFFFF0}.DefineProp("__Item", this.PropertyValueGetter)

    ;https://msdn.microsoft.com/en-us/library/windows/desktop/dd373609(v=vs.85).aspx
    static STATE := {NORMAL:0
        , UNAVAILABLE:0x1
        , SELECTED:0x2
        , FOCUSED:0x4
        , PRESSED:0x8
        , CHECKED:0x10
        , MIXED:0x20
        , INDETERMINATE:0x20
        , READONLY:0x40
        , HOTTRACKED:0x80
        , DEFAULT:0x100
        , EXPANDED:0x200
        , COLLAPSED:0x400
        , BUSY:0x800
        , FLOATING:0x1000
        , MARQUEED:0x2000
        , ANIMATED:0x4000
        , INVISIBLE:0x8000
        , OFFSCREEN:0x10000
        , SIZEABLE:0x20000
        , MOVEABLE:0x40000
        , SELFVOICING:0x80000
        , FOCUSABLE:0x100000
        , SELECTABLE:0x200000
        , LINKED:0x400000
        , TRAVERSED:0x800000
        , MULTISELECTABLE:0x1000000
        , EXTSELECTABLE:0x2000000
        , ALERT_LOW:0x4000000
        , ALERT_MEDIUM:0x8000000
        , ALERT_HIGH:0x10000000
        , PROTECTED:0x20000000
        , VALID:0x7fffffff}.DefineProp("__Item", this.PropertyValueGetter)

    ;https://msdn.microsoft.com/en-us/library/windows/desktop/dd373608(v=vs.85).aspx
    static ROLE := {TITLEBAR:0x1
        , MENUBAR:0x2
        , SCROLLBAR:0x3
        , GRIP:0x4
        , SOUND:0x5
        , CURSOR:0x6
        , CARET:0x7
        , ALERT:0x8
        , WINDOW:0x9
        , CLIENT:0xa
        , MENUPOPUP:0xb
        , MENUITEM:0xc
        , TOOLTIP:0xd
        , APPLICATION:0xe
        , DOCUMENT:0xf
        , PANE:0x10
        , CHART:0x11
        , DIALOG:0x12
        , BORDER:0x13
        , GROUPING:0x14
        , SEPARATOR:0x15
        , TOOLBAR:0x16
        , STATUSBAR:0x17
        , TABLE:0x18
        , COLUMNHEADER:0x19
        , ROWHEADER:0x1a
        , COLUMN:0x1b
        , ROW:0x1c
        , CELL:0x1d
        , LINK:0x1e
        , HELPBALLOON:0x1f
        , CHARACTER:0x20
        , LIST:0x21
        , LISTITEM:0x22
        , OUTLINE:0x23
        , OUTLINEITEM:0x24
        , PAGETAB:0x25
        , PROPERTYPAGE:0x26
        , INDICATOR:0x27
        , GRAPHIC:0x28
        , STATICTEXT:0x29
        , TEXT:0x2a
        , PUSHBUTTON:0x2b
        , CHECKBUTTON:0x2c
        , RADIOBUTTON:0x2d
        , COMBOBOX:0x2e
        , DROPLIST:0x2f
        , PROGRESSBAR:0x30
        , DIAL:0x31
        , HOTKEYFIELD:0x32
        , SLIDER:0x33
        , SPINBUTTON:0x34
        , DIAGRAM:0x35
        , ANIMATION:0x36
        , EQUATION:0x37
        , BUTTONDROPDOWN:0x38
        , BUTTONMENU:0x39
        , BUTTONDROPDOWNGRID:0x3a
        , WHITESPACE:0x3b
        , PAGETABLIST:0x3c
        , CLOCK:0x3d
        , SPLITBUTTON:0x3e
        , IPADDRESS:0x3f
        , OUTLINEBUTTON:0x40}.DefineProp("__Item", this.PropertyValueGetter)

    ;https://msdn.microsoft.com/en-us/library/windows/desktop/dd373600(v=vs.85).aspx
    static NAVDIR := {MIN:0x0
        , UP:0x1
        , DOWN:0x2
        , LEFT:0x3
        , RIGHT:0x4
        , NEXT:0x5
        , PREVIOUS:0x6
        , FIRSTCHILD:0x7
        , LASTCHILD:0x8
        , MAX:0x9}.DefineProp("__Item", this.PropertyValueGetter)

    ;https://msdn.microsoft.com/en-us/library/windows/desktop/dd373634(v=vs.85).aspx
    static SELECTIONFLAG := {NONE:0x0
        , TAKEFOCUS:0x1
        , TAKESELECTION:0x2
        , EXTENDSELECTION:0x4
        , ADDSELECTION:0x8
        , REMOVESELECTION:0x10
        , VALID:0x1f}.DefineProp("__Item", this.PropertyValueGetter)

    ;MSAA Events list:
    ;    https://msdn.microsoft.com/en-us/library/windows/desktop/dd318066(v=vs.85).aspx
    ;What are win events:
    ;    https://msdn.microsoft.com/en-us/library/windows/desktop/dd373868(v=vs.85).aspx
    ;System-Level and Object-level events:
    ;    https://msdn.microsoft.com/en-us/library/windows/desktop/dd373657(v=vs.85).aspx
    ;Console accessibility:
    ;    https://msdn.microsoft.com/en-us/library/ms971319.aspx
    static EVENT := {MIN:0x00000001
        , MAX:0x7FFFFFFF
        , SYSTEM_SOUND:0x0001
        , SYSTEM_ALERT:0x0002
        , SYSTEM_FOREGROUND:0x0003
        , SYSTEM_MENUSTART:0x0004
        , SYSTEM_MENUEND:0x0005
        , SYSTEM_MENUPOPUPSTART:0x0006
        , SYSTEM_MENUPOPUPEND:0x0007
        , SYSTEM_CAPTURESTART:0x0008
        , SYSTEM_CAPTUREEND:0x0009
        , SYSTEM_MOVESIZESTART:0x000A
        , SYSTEM_MOVESIZEEND:0x000B
        , SYSTEM_CONTEXTHELPSTART:0x000C
        , SYSTEM_CONTEXTHELPEND:0x000D
        , SYSTEM_DRAGDROPSTART:0x000E
        , SYSTEM_DRAGDROPEND:0x000F
        , SYSTEM_DIALOGSTART:0x0010
        , SYSTEM_DIALOGEND:0x0011
        , SYSTEM_SCROLLINGSTART:0x0012
        , SYSTEM_SCROLLINGEND:0x0013
        , SYSTEM_SWITCHSTART:0x0014
        , SYSTEM_SWITCHEND:0x0015
        , SYSTEM_MINIMIZESTART:0x0016
        , SYSTEM_MINIMIZEEND:0x0017
        , CONSOLE_CARET:0x4001
        , CONSOLE_UPDATE_REGION:0x4002
        , CONSOLE_UPDATE_SIMPLE:0x4003
        , CONSOLE_UPDATE_SCROLL:0x4004
        , CONSOLE_LAYOUT:0x4005
        , CONSOLE_START_APPLICATION:0x4006
        , CONSOLE_END_APPLICATION:0x4007
        , OBJECT_CREATE:0x8000
        , OBJECT_DESTROY:0x8001
        , OBJECT_SHOW:0x8002
        , OBJECT_HIDE:0x8003
        , OBJECT_REORDER:0x8004
        , OBJECT_FOCUS:0x8005
        , OBJECT_SELECTION:0x8006
        , OBJECT_SELECTIONADD:0x8007
        , OBJECT_SELECTIONREMOVE:0x8008
        , OBJECT_SELECTIONWITHIN:0x8009
        , OBJECT_STATECHANGE:0x800A
        , OBJECT_LOCATIONCHANGE:0x800B
        , OBJECT_NAMECHANGE:0x800C
        , OBJECT_DESCRIPTIONCHANGE:0x800D
        , OBJECT_VALUECHANGE:0x800E
        , OBJECT_PARENTCHANGE:0x800F
        , OBJECT_HELPCHANGE:0x8010
        , OBJECT_DEFACTIONCHANGE:0x8011
        , OBJECT_ACCELERATORCHANGE:0x8012}.DefineProp("__Item", this.PropertyValueGetter)

    class IAccessible {
        __New(oAcc, childId:=0, wId:=0) {
            if ComObjType(oAcc, "Name") != "IAccessible"
                throw Error("Could not access an IAccessible Object")
            this.DefineProp("oAcc", {value:oAcc})
            this.DefineProp("childId", {value:childId})
            this.DefineProp("wId", {value:wId})
        }
        __Get(Name, Params) {
            if !(SubStr(Name,3)="acc") {
                try return this.oAcc.acc%Name%[this.childId]
                try return this.oAcc.acc%Name%(this.childId) ; try method with self
                try return this.oAcc.acc%Name% ; try property
            }
            try return this.oAcc.%Name%[this.childId]
            try return this.oAcc.%Name%(this.childId)
            return this.oAcc.%Name%
        }

        __Item[params*] {
            get {
                oAcc := this
                for _, child in params
                    oAcc := oAcc.GetNthChild(child)
                return oAcc
            }
        }

        __Enum(varCount) {
            maxLen := this.Length, i := 0
            EnumElements(&element) {
                if ++i > maxLen
                    return false
                element := this.GetNthChild(i)
                return true
            }
            EnumIndexAndElements(&index, &element) {
                if ++i > maxLen
                    return false
                index := i
                element := this.GetNthChild(i)
                return true
            }
            return (varCount = 1) ? EnumElements : EnumIndexAndElements
        }

        __Set(Name, Params, Value) {
            if !(SubStr(Name,3)="acc")
                try return this.oAcc.acc%Name%[Params*] := Value
            return this.oAcc.%Name%[Params*] := Value
        }
        
        __Call(Name, Params) {
            if !(SubStr(Name,3)="acc")
                try return this.oAcc.acc%Name%(Params.Length?Params[1]:0)
            return this.oAcc.%Name%(Params*)
        }

        Select(flags) => (this.oAcc.accSelect(flags,this.childId)) ; flags can be any of the SELECTIONFLAG
        DoDefaultAction() => (this.oAcc.accDoDefaultAction(this.childId))
        HitTest(x, y) => (this.oAcc.accHitTest(x, y))

        Name {
            get => (this.oAcc.accName[this.childId])
            set => (this.oAcc.accName[this.childId] := Value)
        } 
        Value {
            get => (this.oAcc.accValue[this.childId])
            set => (this.oAcc.accValue[this.childId] := Value)
        } 
        Role => (this.oAcc.accRole[this.childId])
        RoleText => (Acc.GetRoleText(this.Role))
        Help => (this.oAcc.accHelp[this.childId])
        KeyboardShortcut => (this.oAcc.accKeyboardShortcut[this.childId])
        State => (this.oAcc.accState[this.childId])
        StateText => (Acc.GetStateText(this.oAcc.accState[this.childId]))
        Description => (this.oAcc.accDescription[this.childId])
        DefaultAction => (this.oAcc.accDefaultAction[this.childId])
        Focus => (this.oAcc.accFocus())
        Selection => (this.childId == 0 ? this.oAcc.accState : 0)
        Parent => (Acc.IAccessible(Acc.Query(this.oAcc.accParent),,this.wId))

        IsChild => (this.childId == 0 ? False : True)
        Length => (this.childId == 0 ? this.oAcc.accChildCount : 0)
        Exists {
            get {
                
                if ((state := this.State) == 32768) || (state == 1) || (((pos := this.Location).x==0) && (pos.y==0) && (pos.w==0) && (pos.h==0))
                    return 0
                return 1
            }
        }

        Location {
            get {
                x:=Buffer(4), y:=Buffer(4), w:=Buffer(4), h:=Buffer(4)
                this.oAcc.accLocation(ComValue(0x4003, x.ptr, 1), ComValue(0x4003, y.ptr, 1), ComValue(0x4003, w.ptr, 1), ComValue(0x4003, h.ptr, 1), this.childId)
                Return {x:NumGet(x,0,"int"), y:NumGet(y,0,"int"), w:NumGet(w,0,"int"), h:NumGet(h,0,"int")}
            }
        }

        Children {
            get {
                oAcc := this.oAcc
                cChildren := oAcc.accChildCount, Children := Array()
                varChildren := Buffer(cChildren * (8+2*A_PtrSize))
                try {
                    if DllCall("oleacc\AccessibleChildren", "ptr", ComObjValue(oAcc), "int",0, "int", cChildren, "ptr", varChildren, "int*", cChildren) > -1 {
                        Loop cChildren {
                            i := (A_Index-1) * (A_PtrSize * 2 + 8) + 8
                            child := NumGet(varChildren, i, "ptr")
                            Children.Push(NumGet(varChildren, i-8, "ptr") = 9 ? Acc.IAccessible(Acc.Query(child),,this.wId) : Acc.IAccessible(oAcc, child, this.wId))
                            NumGet(varChildren, i-8, "ptr") = 9 ? ObjRelease(child) : ""
                        }
                        Return Children
                    }
                }
                throw Error("AccessibleChildren DllCall Failed", -1)
            }
        }
    
        GetNthChild(n) {
            cChildren := this.oAcc.accChildCount
            if n < 1 || n > cChildren
                throw IndexError("Child index " n " is out of bounds", -1)
            varChildren := Buffer(cChildren * (8+2*A_PtrSize))
            try {
                if DllCall("oleacc\AccessibleChildren", "ptr",ComObjValue(this.oAcc), "int",0, "int",cChildren, "ptr",varChildren, "int*",cChildren) > -1 {
                    i := (n-1) * (A_PtrSize * 2 + 8) + 8
                    child := NumGet(varChildren, i, "ptr")
                    oChild := NumGet(varChildren, i-8, "ptr") = 9 ? Acc.IAccessible(Acc.Query(child),,this.wId) : Acc.IAccessible(this.oAcc, child, this.wId)
                    NumGet(varChildren, i-8, "ptr") = 9 ? ObjRelease(child) : ""
                    Return oChild
                }
            }
            throw Error("AccessibleChildren DllCall Failed", -1)
        }

    	; Returns an object containing the x, y coordinates and width and height: {x:x coordinate, y:y coordinate, w:width, h:height}. relativeTo can be client, window or screen, default is A_CoordModeMouse.
        GetLocation(relativeTo:="") { 
            relativeTo := (relativeTo == "") ? A_CoordModeMouse : relativeTo, loc := this.Location
            if (relativeTo = "screen")
                return loc
            else if (relativeTo = "window") {
                RECT := Buffer(16)
                DllCall("user32\GetWindowRect", "Int", this.wId, "Ptr", RECT)
                return {x:(loc.x-NumGet(RECT, 0, "Int")), y:(loc.y-NumGet(RECT, 4, "Int")), w:loc.w, h:loc.h}
            } else if (relativeTo = "client") {
                pt := Buffer(8), NumPut("int",loc.x,pt), NumPut("int",loc.y,pt,4)
                DllCall("ScreenToClient", "Int", this.wId, "Ptr", pt)
                return {x:NumGet(pt,0,"int"), y:NumGet(pt,4,"int"), w:loc.w, h:loc.h}
            } else
                throw Error(relativeTo "is not a valid CoordMode",-1)
        }

        IsEqual(oCompare) {
            try loc1 := this.Location, loc2 := oCompare.Location
            catch 
                return 0
            if (loc1.x != loc2.x) || (loc1.y != loc2.y) || (loc1.w != loc2.w) || (loc1.h != loc2.h) || (this.Role != oCompare.Role)
                return 0
            return 1
        }

        ; Finds the first element matching the condition (see description under ValidateCondition)
        ; Scope is the search scope: 1=element itself; 2=direct children; 4=descendants (including children of children)
        FindFirst(condition, scope:=4) {
            if scope&1 {
                if this.ValidateCondition(condition)
                    return this
            }
            if scope>1 {
                for _, el in this {
                    if el.ValidateCondition(condition)
                        return el
                    else if scope&4 {
                        try return el.FindFirst(condition, scope ^= 1)
                    } 
                }
            }
            throw Error("Matching Acc object not found")
        }
        ; Returns an array of elements matching the condition (see description under ValidateCondition)
        FindAll(condition, scope:=4) {
            RecursiveFind(this, condition, scope, &matches := [])
            return matches

            RecursiveFind(element, condition, scope, &matches) {
                if scope&1 {
                    if element.ValidateCondition(condition)
                        matches.Push(element)
                }
                if scope>1 {
                    for _, el in element {
                        if el.ValidateCondition(condition)
                            matches.Push(el)
                        if scope&4
                            RecursiveFind(el, condition, scope ^= 1, &matches)
                    }
                }
            }          
        }
        WaitElementExist(conditionOrPath, scope:=4, timeOut:=-1) {
            startTime := A_TickCount
            while ((timeOut < 1) ? 1 : (A_tickCount - startTime < timeOut)) {
                try {
                    if IsObject(conditionOrPath)
                        return this.FindFirst(conditionOrPath, scope)
                    else
                        return this[StrSplit(conditionOrPath,",")*]
                }
                Sleep 40
            }
            throw Error("WaitElementExist timed out")
        }

        /*
            Checks whether the element matches a provided condition.
            Everything inside {} is an "and" condition
            Everything inside [] is an "or" condition
            Object key "not" creates a not condition

            matchmode key defines the MatchMode: 1=must start with; 2=can contain anywhere in string; 3=exact match; RegEx

            casesensitive key defines case sensitivity: True=case sensitive; False=case insensitive

            {Name:"Something"} => Name must match "Something" (case sensitive)
            {Name:"Something", RoleText:"something else"} => Name must match "Something" and RoleText must match "something else"
            [{Name:"Something", Role:42}, {Name:"Something2", RoleText:"something else"}] => Name=="Something" and Role==42 OR Name=="Something2" and RoleText=="something else"
            {Name:"Something", not:[RoleText:"something", RoleText:"something else"]} => Name must match "something" and RoleText cannot match "something" nor "something else"
        */
        ValidateCondition(oCond) {
            if Type(oCond) = "Array" { ; or condition
                for _, c in oCond
                    if this.ValidateCondition(c)
                        return 1
                return 0
            }
            matchmode := 3, casesensitive := 1, notCond := False
            
            if oCond.HasOwnProp("matchmode")
                matchmode := oCond.matchmode
            if oCond.HasOwnProp("casesensitive")
                casesensitive := oCond.casesensitive
            for prop, cond in oCond.OwnProps() {
                switch Type(cond) { ; and condition
                    case "String", "Integer":
                        if prop ~= "casesensitive|matchmode"
                            continue
                        propValue := ""
                        try propValue := this.%prop%
                        switch matchmode, 0 {
                            case 2:
                                if !InStr(propValue, cond, casesensitive)
                                    return 0
                            case 1:
                                if !((casesensitive && (SubStr(propValue, 1, StrLen(cond)) == cond)) || (!casesensitive && (SubStr(propValue, 1, StrLen(cond)) = cond)))
                                    return 0
                            case "Regex":
                                if !(propValue ~= cond)
                                    return 0
                            default:
                                if !((casesensitive && (propValue == cond)) || (!casesensitive && (propValue = cond)))
                                    return 0
                        }
                    default:
                        if (prop = "not") ? this.ValidateCondition(cond) : !this.ValidateCondition(cond)
                            return 0
                }
            }
            return 1
        }

        ; Outputs relevant information about the element
        Dump() {
            Value := "", Name := "", StateText := "", State := "", DefaultAction := "", Description := "", KeyboardShortcut := "", Help := "", Pos := {x:0,y:0,w:0,h:0}
            for _, v in ["Value", "Name", "StateText", "State", "DefaultAction", "Description", "KeyboardShortcut", "Help"]
                try %v% := this.%v%
            try Pos := this.Location
            return "RoleText: " Acc.GetRoleText(Role := this.Role) " Role: " Role " [Location: {x:" Pos.x ",y:" Pos.y ",w:" Pos.w ",h:" Pos.h "}]" " [Name: " (Name ?? "") "] [Value: " (Value ?? "")  "]" (StateText ? " [StateText: " StateText "]" : "") (State ? " [State: " State "]" : "") (DefaultAction ? " [DefaultAction: " DefaultAction "]" : "") (Description ? " [Description: " Description "]" : "") (KeyboardShortcut ? " [KeyboardShortcut: " KeyboardShortcut "]" : "") (Help ? " [Help: " Help "]" : "") (this.childId ? " ChildId: " this.childId : "")
        }

        DumpAll() {
            return RecurseTree(this, this.Dump() "`n")
            RecurseTree(oAcc, tree, path:="") {
                try {
                    if !oAcc.Length
                        return tree
                } catch
                    return tree
                
                For i, oChild in oAcc.Children {
                    tree .= path (path?",":"") i ": " oChild.Dump() "`n"
                    tree := RecurseTree(oChild, tree, path (path?",":"") i)
                }
                return tree
            }
        }
        /*
            Highlights the element for a chosen period of time
            Possible showTime values:
                Unset: removes the highlighting
                0: Indefinite highlighting
                Positive integer (eg 2000): will highlight and pause for the specified amount of time in ms
                Negative integer: will highlight for the specified amount of time in ms, but script execution will continue
        */
        Highlight(showTime:=unset, color:="Red", d:=2) {
            static range := [], removeHighlight := ObjBindMethod(Acc.IAccessible, "Highlight")
            if !IsSet(showTime) {
                for _, r in range
                    r.Destroy()
                range := []
                return
            }
            loc := this.Location
            if !IsObject(loc)
                return
            Loop 4 {
                range.Push(Gui("+AlwaysOnTop -Caption +ToolWindow -DPIScale +E0x08000000"))
            }
            Loop 4
            {
                i:=A_Index
                , x1:=(i=2 ? loc.x+loc.w : loc.x-d)
                , y1:=(i=3 ? loc.y+loc.h : loc.y-d)
                , w1:=(i=1 or i=3 ? loc.w+2*d : d)
                , h1:=(i=2 or i=4 ? loc.h+2*d : d)
                range[i].BackColor := color
                range[i].Show("NA x" . x1 . " y" . y1 . " w" . w1 . " h" . h1)
            }
            if showTime > 0 {
                Sleep(showTime)
                this.Highlight()
            } else if showTime < 0
                SetTimer(removeHighlight, -Abs(showTime))
        }


        ; Click the center of the element.
        ; If WhichButton is a number, then Sleep will be called with that number. Eg Click(200) will sleep 200ms after clicking
        ; If ClickCount is a number >=10, then Sleep will be called with that number. To click 10+ times and sleep after, specify "ClickCount SleepTime". Ex: Click("left", 200) will sleep 200ms after clicking. Ex: Click("left", "20 200") will left-click 20 times and then sleep 200ms.
        ; If Relative is "Rel" or "Relative" then X and Y coordinates are treated as offsets from the current mouse position. Otherwise it expects offset values for both X and Y (eg "-5 10" would offset X by -5 and Y by +10).
        Click(WhichButton:="left", ClickCount:=1, DownOrUp:="", Relative:="") {		
            rel := [0,0], pos := this.GetLocation()
            if (Relative && !InStr(Relative, "rel"))
                rel := StrSplit(Relative, " "), Relative := ""
            cCount := 1, SleepTime := -1
            if IsInteger(WhichButton)
                SleepTime := WhichButton, WhichButton := "left"
            if !IsInteger(ClickCount) && InStr(ClickCount, " ") {
                sCount := StrSplit(ClickCount, " ")
                cCount := sCount[1], SleepTime := sCount[2]
            } else if ClickCount > 9 {
                SleepTime := cCount, cCount := 1
            }
            Click((pos.x+pos.w//2+rel[1]) " " (pos.y+pos.h//2+rel[2]) " " WhichButton (ClickCount ? " " ClickCount : "") (DownOrUp ? " " DownOrUp : "") (Relative ? " " Relative : ""))
            Sleep(SleepTime)
        }

        ; ControlClicks the element after getting relative coordinates with GetLocation("client"). 
        ; If WhichButton is a number, then a Sleep will be called afterwards. Ex: ControlClick(200) will sleep 200ms after clicking. Same for ControlClick("ahk_id 12345", 200)
        ControlClick(WhichButton:="left", ClickCount:=1, Options:="") { 
            pos := this.GetLocation("client")
            ControlClick("X" pos.x+pos.w//2 " Y" pos.y+pos.h//2, this.wId,, IsInteger(WhichButton) ? "left" : WhichButton, ClickCount, Options)
            if IsInteger(WhichButton)
                Sleep(WhichButton)
        }
    }

    static ObjectFromPoint(x:=unset, y:=unset, &idChild := "") {
        if !(IsSet(x) && IsSet(y))
            DllCall("GetCursorPos", "int64P", &pt64:=0), x := 0xFFFFFFFF & pt64, y := pt64 >> 32
        else {
            pt64 := y << 32 | x
        }
        pvarChild := Buffer(8 + 2 * A_PtrSize)
        if DllCall("oleacc\AccessibleObjectFromPoint", "int64",pt64, "ptr*",&ppAcc := 0, "ptr",pvarChild) = 0
        {	; returns a pointer from which we get a Com Object
            idChild:=NumGet(pvarChild,8,"UInt")
            wId := DllCall("GetAncestor", "UInt", DllCall("user32.dll\WindowFromPoint", "int64",  pt64), "UInt", GA_ROOT := 2) ; hwnd from point by SKAN
            return Acc.IAccessible(ComValue(9, ppAcc), idChild, wId)
        }
    }
    
    static ObjectFromWindow(hWnd:="A", idObject := 0) {
        if !hWnd
            throw Error("Invalid window handle provided", -2)
        if !IsInteger(hWnd)
            hWnd := WinExist(hWnd)
        IID := Buffer(16)
        if DllCall("oleacc\AccessibleObjectFromWindow", "ptr",hWnd, "uint",idObject &= 0xFFFFFFFF
                , "ptr",-16 + NumPut("int64", idObject == 0xFFFFFFF0 ? 0x46000000000000C0 : 0x719B3800AA000C81, NumPut("int64", idObject == 0xFFFFFFF0 ? 0x0000000000020400 : 0x11CF3C3D618736E0, IID))
                , "ptr*", ComObj := ComValue(9,0)) = 0
            Return Acc.IAccessible(ComObj,,hWnd)
    }
    
    static ObjectFromPath(ChildPath, hWnd:="A") {
        oAcc := Acc.ObjectFromWindow(hWnd)
        ChildPath := StrReplace(StrReplace(ChildPath, ".", ","), " ")
        Loop Parse ChildPath, ","
        {
            oAcc := oAcc.GetNthChild(A_LoopField)
        }
        Return oAcc
    }

    static ObjectFromEvent(hWnd, idObject, idChild) {
        if (DllCall("oleacc\AccessibleObjectFromEvent"
                , "Ptr", hWnd
                , "UInt", idObject
                , "UInt", idChild
                , "Ptr*", pacc := ComValue(9,0)
                , "Ptr", varChild := Buffer(16)) = 0) {
            return Acc.IAccessible(pacc, NumGet(varChild, 8, "UInt"), hWnd)
        }
        throw Error("ObjectFromEvent failed", -1)
    }

    static GetRootElement() {
        return Acc.ObjectFromWindow(0x10010)
    }
       
    static Query(pAcc) {
        oCom := ComObjQuery(pAcc, "{618736e0-3c3d-11cf-810c-00aa00389b71}")
        ObjAddRef(oCom.ptr)
        Try Return ComValue(9, oCom.ptr)
    }

    static GetRoleText(nRole) {
        if !IsInteger(nRole)
            throw TypeError("The specified role is not an integer!",-2)
        nRole := Integer(nRole)
        nSize := DllCall("oleacc\GetRoleText", "Uint", nRole, "Ptr", 0, "Uint", 0)
        VarSetStrCapacity(&sRole, nSize+2)
        DllCall("oleacc\GetRoleText", "Uint", nRole, "str", sRole, "Uint", nSize+2)
        Return sRole
    }

    static GetStateText(nState) {
        nSize := DllCall("oleacc\GetStateText"
          , "Uint"	, nState
          , "Ptr" 	, 0
          , "Uint"	, 0)
        VarSetStrCapacity(&sState, nSize+2)
        DllCall("oleacc\GetStateText"
          , "Uint"	, nState
          , "str" 	, sState
          , "Uint"	, nSize+2)
        return sState
    }

    static RegisterWinEvent(event, callback) {
        pCallback := CallbackCreate(this.GetMethod("HandleWinEvent").Bind(this, callback), "F", 7)
        hook := Acc.SetWinEventHook(event,event,pCallback)
        return {__Hook:hook, __Callback:pCallback, __Delete:{ call: (*) => (this.UnhookWinEvent(hook), CallbackFree(pCallback)) }}
    }

    static HandleWinEvent(fCallback, hWinEventHook, Event, hWnd, idObject, idChild, dwEventThread, dwmsEventTime) {
        return fCallback(Acc.ObjectFromEvent(hWnd, idObject, idChild), Event, dwmsEventTime&0x7FFFFFFF)
    }

    static SetWinEventHook(eventMin, eventMax, pCallback) {
        DllCall("ole32\CoInitialize", "Uint", 0)
        Return DllCall("SetWinEventHook", "Uint", eventMin, "Uint", eventMax, "Uint", 0, "UInt", pCallback, "Uint", 0, "Uint", 0, "Uint", 0)
    }

    static UnhookWinEvent(hHook) {
        Return DllCall("UnhookWinEvent", "Ptr", hHook)
    }

    class Viewer {
        __New() {
            this.Stored := {mwId:0, Trees:Map(), TreeView:Map()}
            this.Capturing := False
            this.gViewer := Gui("AlwaysOnTop Resize","AccViewer")
            this.gViewer.OnEvent("Close", (*) => ExitApp())
            this.gViewer.OnEvent("Size", this.GetMethod("gViewer_Size").Bind(this))
            this.gViewer.Add("Text", "w100", "Window Info").SetFont("bold")
            this.LVWin := this.gViewer.Add("ListView", "h140 w250", ["Property", "Value"])
            this.LVWin.OnEvent("ContextMenu", LV_CopyTextMethod := this.GetMethod("LV_CopyText").Bind(this))
            this.LVWin.ModifyCol(1,100)
            this.LVWin.ModifyCol(2,140)
            for _, v in ["Title", "Text", "Id", "Location", "Class(NN)", "Process", "PID"]
                this.LVWin.Add(,v,"")
            this.gViewer.Add("Text", "w100", "Acc Info").SetFont("bold")
            this.LVProps := this.gViewer.Add("ListView", "h220 w250", ["Property", "Value"])
            this.LVProps.OnEvent("ContextMenu", LV_CopyTextMethod)
            this.LVProps.ModifyCol(1,100)
            this.LVProps.ModifyCol(2,140)
            for _, v in ["RoleText", "Role", "Value", "Name", "Location", "StateText", "State", "DefaultAction", "Description", "KeyboardShortcut", "Help", "ChildId"]
                this.LVProps.Add(,v,"")
            this.ButCapture := this.gViewer.Add("Button", "xp+60 y+10 w130", "Start capturing (F1)")
            this.ButCapture.OnEvent("Click", this.CaptureHotkeyFunc := this.GetMethod("ButCapture_Click").Bind(this))
            HotKey("~F1", this.CaptureHotkeyFunc)
            this.SBMain := this.gViewer.Add("StatusBar",, "  Start capturing, then hold cursor still to construct tree")
            this.SBMain.OnEvent("Click", this.GetMethod("SBMain_Click").Bind(this))
            this.SBMain.OnEvent("ContextMenu", this.GetMethod("SBMain_Click").Bind(this))
            this.gViewer.Add("Text", "x278 y10 w100", "Acc Tree").SetFont("bold")
            this.TVAcc := this.gViewer.Add("TreeView", "x275 y25 w250 h390")
            this.TVAcc.OnEvent("Click", this.GetMethod("TVAcc_Click").Bind(this))
            this.TVAcc.Add("Start capturing to show tree")
            this.gViewer.Show()
        }
        gViewer_Size(GuiObj, MinMax, Width, Height) {
            this.TVAcc.GetPos(&TVAccX, &TVAccY, &TVAccWidth, &TVAccHeight)
            this.TVAcc.Move(,,Width-TVAccX-10,Height-TVAccY-25)
            this.TVAcc.GetPos(&LVPropsX, &LVPropsY, &LVPropsWidth, &LVPropsHeight)
            this.LVProps.Move(,,,Height-LVPropsY-225)
            this.ButCapture.Move(,Height -50)
        }
        ButCapture_Click(GuiCtrlObj?, Info?) {
            if this.Capturing {
                this.StopCapture()
                return
            }
            this.Capturing := True
            HotKey("~F1", this.CaptureHotkeyFunc, "Off")
            HotKey("~Esc", this.CaptureHotkeyFunc)
            this.TVAcc.Delete()
            this.TVAcc.Add("Hold cursor still to construct tree")
            this.ButCapture.Text := "Stop capturing (Esc)"
            this.CaptureCallback := this.GetMethod("CaptureCycle").Bind(this)
            SetTimer(this.CaptureCallback, 200)
        }
        LV_CopyText(GuiCtrlObj, Info, *) {
            ToolTip("Copied: " (A_Clipboard := GuiCtrlObj.GetText(Info,2)))
            SetTimer((*) => ToolTip(), -3000)
        }
        SBMain_Click(GuiCtrlObj, Info, *) {
            if InStr(this.SBMain.Text, "Path:") {
                ToolTip("Copied: " (A_Clipboard := SubStr(this.SBMain.Text, 9)))
                SetTimer((*) => ToolTip(), -3000)
            }
        }
        StopCapture(GuiCtrlObj:=0, Info:=0) {
            if this.Capturing {
                this.Capturing := False
                this.ButCapture.Text := "Start capturing (F1)"
                HotKey("~Esc", this.CaptureHotkeyFunc, "Off")
                HotKey("~F1", this.CaptureHotkeyFunc)
                SetTimer(this.CaptureCallback, 0)
                this.Stored.oAcc.Highlight()
                return
            }
        }
        CaptureCycle() {
            MouseGetPos(&mX, &mY, &mwId)
            oAcc := Acc.ObjectFromPoint()
            if this.Stored.HasOwnProp("oAcc") && oAcc.IsEqual(this.Stored.oAcc) {
                if this.FoundTime != 0 && ((A_TickCount - this.FoundTime) > 1000) {
                    if (mX == this.Stored.mX) && (mY == this.Stored.mY) 
                        this.ConstructTreeView(), this.FoundTime := 0
                    else 
                        this.FoundTime := A_TickCount
                }
                this.Stored.mX := mX, this.Stored.mY := mY
                return
            }
            this.LVWin.Delete()
            WinGetPos(&mwX, &mwY, &mwW, &mwH, mwId)
            propsOrder := ["Title", "Text", "Id", "Location", "Class(NN)", "Process", "PID"]
            props := Map("Title", WinGetTitle(mwId), "Text", WinGetText(mwId), "Id", mwId, "Location", "x: " mwX " y: " mwY " w: " mwW " h: " mwH, "Class(NN)", WinGetClass(mwId), "Process", WinGetProcessName(mwId), "PID", WinGetPID(mwId))
            for propName in propsOrder
                this.LVWin.Add(,propName,props[propName])
            this.LVProps_Populate(oAcc)
            this.Stored.mwId := mwId, this.Stored.oAcc := oAcc, this.Stored.mX := mX, this.Stored.mY := mY, this.FoundTime := A_TickCount
        }
        LVProps_Populate(oAcc) {
            oAcc.Highlight() ; Clear
            oAcc.Highlight(0) ; Indefinite show
            this.LVProps.Delete()
            Location := {x:0,y:0,w:0,h:0}, RoleText := "", Role := "", Value := "", Name := "", StateText := "", State := "", DefaultAction := "", Description := "", KeyboardShortcut := "", Help := "", ChildId := ""
            for _, v in ["RoleText", "Role", "Value", "Name", "Location", "StateText", "State", "DefaultAction", "Description", "KeyboardShortcut", "Help", "ChildId"] {
                try %v% := oAcc.%v%
                this.LVProps.Add(,v, v = "Location" ? ("x: " %v%.x " y: " %v%.y " w: " %v%.w " h: " %v%.h) :%v%)
            }
        }
        TVAcc_Click(GuiCtrlObj, Info) {
            if this.Capturing
                return
            oAcc := this.Stored.TreeView[Info]
            this.LVProps_Populate(oAcc)
        }
        ConstructTreeView() {
            this.TVAcc.Delete()
            this.TVAcc.Opt("-Redraw")
            this.SBMain.SetText("  Path: " this.GetAccPathTopDown(this.Stored.mwId, this.Stored.oAcc))
            this.RecurseTreeView(Acc.ObjectFromWindow(this.Stored.mwId))
            this.TVAcc.Opt("+Redraw")
            for k, v in this.Stored.TreeView
                if this.Stored.oAcc.IsEqual(v)
                    this.TVAcc.Modify(k, "Vis Select")
            

        }
        RecurseTreeView(oAcc, parent:="") {
            try elDesc := " `"" oAcc.Name "`""
            catch
                elDesc := " `"`""
            elDesc := oAcc.RoleText elDesc
            this.Stored.TreeView[TWEl := this.TVAcc.Add(elDesc, parent, "Expand")] := oAcc
            for k, v in oAcc
                this.RecurseTreeView(v, TWEl)
        }
        BuildAccTreeRecursive(oAcc, tree, path:="") {
            if !IsObject(oAcc) || !oAcc.length
                return tree
            For i, oChild in oAcc {
                tree[path (path?",":"") i] := oChild
                tree := this.BuildAccTreeRecursive(oChild, tree, path (path?",":"") i)
            }
            return tree
        }
        GetAccPathTopDown(hwnd, oAcc, updateTree:=False) {
            if !this.Stored.Trees.Has(hwnd) || updateTree
                this.Stored.Trees[hwnd] := this.BuildAccTreeRecursive(Acc.ObjectFromWindow(hwnd), Map("Root", oAcc))
            for k, v in this.Stored.Trees[hwnd] {
                if v.IsEqual(oAcc)
                    return k
            }
        }
    }
}

