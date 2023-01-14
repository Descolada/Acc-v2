# Acc-v2
Acc library for AHK v2


## Notable changes (compared to Acc v1):
Acc v2 in not a port of AHK v1 Acc library, but instead a complete redesign to incorporate more object-oriented approaches. 

    1) All Acc elements are now array-like objects, where the "Length" property contains the number of children, 
       any nth children can be accessed with element[n], and children can be iterated over with for loops.
    2) Acc main functions are contained in the global Acc class/variable
    3) Element methods are contained inside element objects
    4) Element properties can be used without the "acc" prefix
    5) ChildIds have been removed (are handled in the backend), but can be accessed through el.ChildId
    6) Additional methods have been added for elements, such as FindElement, FindElements, Click
    7) Acc constants are included in the Acc object
    8) AccViewer is built into the library: when ran directly the AccViewer will show, when included
       in another script then it won't show (but can be accessed by calling Acc.Viewer())
    9) Some main Acc methods have been given aliases: eg ObjectFromWindow -> ElementFromHandle.

## Short introduction
Acc (otherwise known as IAccessible or MSAA) is a library to get information about (and sometimes interact with) windows, controls, and window elements that are otherwise not accessible with AHKs Control functions. 

When Acc.ahk is ran alone, it displays the AccViewer: a window inspecter to get Acc information from elements. In the AccViewer on the left side are displayed all the available properties for elements (Value, Name etc), and on the right side will be displayed the Acc tree which shows how elements are related to eachother. In the bottom is displayed the Acc path (right-click to copy it), which can be used to get that specific element with the Acc library.  

To get started with Acc, first [include](https://lexikos.github.io/v2/docs/commands/_Include.htm) it in your program with [#include](https://lexikos.github.io/v2/docs/commands/_Include.htm). If Acc.ahk is in the same folder as your script use `#include Acc.ahk`, if it is in the Lib folder then use `#include <Acc>`.  

All main Acc properties and methods are accessible through the global variable `Acc`, which is created when Acc.ahk is included in your script. 
To access Acc elements, first you need to get a starting point element (usually a window) and save it in a variable: something like `oAcc := Acc.ElementFromHandle(WinTitle)`. WinTitle uses the same rules as any other AHK function, so SetTitleMatchMode also applies to it.  

To get elements from the window element, you can use the Acc path from AccViewer: for example `oEl := oAcc[4,1,4]` would get the windows 4th sub-element, then the sub-elements 1st child, and then its 4th child.  

All the properties displayed in AccViewer can be accessed from the element: `oEl.Name` will get that elements' name or throw an error if the name doesn't exist. Most properties are read-only, but the Value property can sometimes be changed with `oEl.Value := "newvalue"`  

Element methods can be used in the same way. To do the default action (usually clicking), use `oEl.DoDefaultAction()`, to highlight the element for 2 seconds use `oEl.Highlight(2000)`. Dump info about a specific element with `oEl.Dump()` and dump info about sub-elements as well with `oEl.DumpAll()` (to get paths and info about all elements in a window, use it on the window element: `MsgBox( Acc.ElementFromHandle(WinTitle).DumpAll() )`  

Some examples of how Acc.ahk can be used are included in the Examples folder.  

# Acc methods

    ElementFromPoint(x:=unset, y:=unset, &idChild := "", activateChromium := True)
        Gets an Acc element from screen coordinates X and Y (NOT relative to the active window).
    ElementFromHandle(hWnd:="A", idObject := 0, activateChromium := True)
        Gets an Acc element from a WinTitle, by default the active window. This can also be a Control handle.
        Additionally idObject can be specified from Acc.ObjId constants (eg to get the Caret location).
    GetRootElement()
        Gets the Acc element for the Desktop
    ActivateChromiumAccessibility(hWnd) 
        Sends the WM_GETOBJECT message to the Chromium document element and waits for the 
        app to be accessible to Acc. This is called when ObjectFromPoint or ObjectFromWindow 
        activateChromium flag is set to True. A small performance increase may be gotten 
        if that flag is set to False when it is not needed.
    RegisterWinEvent(event, callback, PID:=0) 
    RegisterWinEvent(eventMin, eventMax, callback, PID:=0)
        Registers an event or event range from Acc.Event to a callback function and returns
            a new object containing the WinEventHook
        EventMax is an optional variable: if only eventMin and callback are provided, then
            only that single event is registered. If all three arguments are provided, then
            an event range from eventMin to eventMax are registered to the callback function.
        The callback function needs to have two arguments: 
            CallbackFunction(oAcc, EventInfo)

            When the callback function is called:
            oAcc will be the Acc element that called the event
            EventInfo will be an object containing the following properties: 
                Event - an Acc.Event constant
                EventTime - when the event was triggered in system time
                WinID - handle of the window that sent the event 
                ControlID - handle of the control that sent the event, which depending on the
                    window will be the window itself or a control
                ObjId - the object Id (Acc.ObjId) the event was called with
        PID is the Process ID of the process/window the events will be registered from. By default
            events from all windows are registered.
        Unhooking of the event handler will happen once the returned object is destroyed
        (either when overwritten by a constant, or when the script closes).

    Legacy methods:
    SetWinEventHook(eventMin, eventMax, pCallback)
    UnhookWinEvent(hHook)
    ElementFromPath(ChildPath, hWnd:="A")
        Same as ElementFromHandle[comma-separated path]
    GetRoleText(nRole)
        Same as element.RoleText
    GetStateText(nState)
        Same as element.StateText
    Query(pAcc)
        For internal use
        
# Acc constants

    Constants can be accessed as properties (eg Acc.OBJID.CARET), or the property name can be
      accessed by getting as an item (eg Acc.OBJID[0xFFFFFFF8])

    ObjId - object identifiers that identify categories of accessible objects within a window. 
    State - used to describe the state of objects in an application UI. These are returned by Element.State or Element.StateText.
    Role - used to describe the roles of various UI objects in an application. These are returned by Element.Role or Element.RoleText.
    NavDir - indicate the spatial (up, down, left, and right) or logical (first child, 
        last, next, and previous) direction used with Element.Navigate() to navigate from one 
        user interface element to another within the same container.
    SelectionFlag - used to specify how an accessible object becomes selected or takes the focus.
        These are used by Element.Select().
    Event - events that are generated by the operating system and by applications. These are
        used when dealing with RegisterWinEvent.

More thorough explanations for the constants are available [in Microsoft documentations](https://docs.microsoft.com/en-us/windows/win32/winauto/constants-and-enumerated-types).

# Acc element properties
    Element[n]          => Gets the nth element. Multiple of these can be used like a path:
                                Element[4,1,4] will select 4th childs 1st childs 4th child
                            Conditions (see ValidateCondition) are supported: 
                                Element[4,{Name:"Something"}] will select the fourth childs first child matching the name "Something"
                            Conditions also accept an index (or i) parameter to select from multiple similar elements
                                Element[{Name:"Something", i:3}] selects the third element of elements with name "Something"
                            Negative index will select from the last element
                                Element[{Name:"Something", i:-1}] selects the last element of elements with name "Something"
    Name                => Gets or sets the name. All objects support getting this property.
    Value               => Gets or sets the value. Not all objects have a value.
    Role                => Gets the Role of the specified object in integer form. All objects support this property.
    RoleText            => Role converted into text form. All objects support this property.
    Help                => Retrieves the Help property string of an object. Not all objects support this property.
    KeyboardShortcut    => Retrieves the specified object's shortcut key or access key. Not all objects support this property.
    State               => Retrieves the current state in integer form. All objects support this property.
    StateText           => State converted into text form
    Description         => Retrieves a string that describes the visual appearance of the specified object. Not all objects have a description.
    DefaultAction       => Retrieves a string that indicates the object's default action. Not all objects have a default action.
    Focus               => Returns the focused child element (or itself).
                               If no child is focused, an error is thrown
    Selection           => Retrieves the selected children of this object. All objects that support selection must support this property.
    Parent              => Returns the parent element. All objects support this property.
    IsChild             => Checks whether the current element is of child type (this is usually not needed)
    Length              => Returns the number of children the element has
    Location            => Returns the object's current screen location in an object {x,y,w,h}
    Children            => Returns all children as an array (usually not required)
    Exists              => Checks whether the element is still alive and accessible
    ControlID           => ID (hwnd) of the control associated with the element
    WinID               => ID of the window the element belongs to
    accessible          => ComObject of the underlying IAccessible (this is usually not needed)
    childId             => childId of the underlying IAccessible (this is usually not needed)
    
# Acc element methods

    Select(flags)
        Modifies the selection or moves the keyboard focus of the specified object. flags can be any of the Acc.SelectionFlag constants
    DoDefaultAction()
        Performs the specified object's default action. Not all objects have a default action.
    GetNthChild(n)
        This is equal to oElement[n]
    GetLocation(relativeTo:="")
        Returns an object containing the x, y coordinates and width and height: {x:x coordinate, y:y coordinate, w:width, h:height}. 
        relativeTo can be client, window or screen, default is A_CoordModeMouse.
    IsEqual(oCompare)
        Checks whether the element is equal to another element (oCompare)
    FindElement(condition, scope:=4, index:=1, order:=0, depth:=-1) 
        Condition: A condition object (see ValidateCondition). This condition object can also contain named argument values:
            FindElement({Name:"Something", scope:"Subtree"})
        Scope: the search scope (Acc.SCOPE value): Element, Children, Family (Element+Children), Descendants, SubTree (Element+Descendants). Default is Descendants.
        Index: can be used to search for i-th element. 
            Like the other parameters, this can also be supplied in the condition with index or i:
                FindElement({Name:"Something", i:3}) finds the third element with name "Something"
            Negative index reverses the search direction:
                FindElement({Name:"Something", i:-1}) finds the last element with name "Something"
            Since index/i needs to be a key-value pair, then to use it with an "or" condition
            it must be inside an object ("and" condition), for example with key "or":
                FindElement({or:[{Name:"Something"}, {Name:"Something else"}], index:2})
        Order: defines the order of tree traversal (Acc.TreeTraversalOptions value): 
            Default, LastToFirst, PostOrder. Default is FirstToLast and PreOrder.
    FindElements(condition:=True, scope:=4, depth:=-1)
        Returns an array of elements matching the condition (see description under ValidateCondition)
        The returned elements also have the "Path" property with the found elements path
    WaitElement(conditionOrPath, timeOut:=-1, scope:=4, index:=1, order:=0, depth:=-1)
        Waits an element to be detectable in the Acc tree. This doesn't mean that the element
        is visible or interactable, use WaitElementExist for that. 
        Timeout less than 1 waits indefinitely, otherwise is the wait time in milliseconds
        A timeout returns 0.
    WaitElementExist(conditionOrPath, timeOut:=-1, scope:=4, index:=1, order:=0, depth:=-1)
        Waits an element exist that matches a condition or a path. 
        Timeout less than 1 waits indefinitely, otherwise is the wait time in milliseconds
        A timeout returns 0.
    Normalize(condition)
        Checks whether the current element or any of its ancestors match the condition, 
        and returns that element. If no element is found, an error is thrown.
    ValidateCondition(condition)
        Checks whether the element matches a provided condition.
        Everything inside {} is an "and" condition, or a singular condition with options
        Everything inside [] is an "or" condition
        "not" key creates a not condition
        "matchmode" key (short form: "mm") defines the MatchMode: StartsWith, Substring, Exact, RegEx (Acc.MATCHMODE values)
        "casesensitive" key (short form: "cs") defines case sensitivity: True=case sensitive; False=case insensitive
        Any other key (but recommended is "or") can be used to use "or" condition inside "and" condition.
        Additionally, when matching for location then partial matching can be used (eg only width and height)
            and relative mode (client, window, screen) can be specified with "relative" or "r".
        An empty object {} is used as "unset" or "N/A" value.

        For methods which use this condition, it can also contain named arguments:
            oAcc.FindElement({Name:"Something", scope:"Subtree", order:"LastToFirst"})
            is equivalent to FindElement({Name:"Something"}, "Subtree",, "LastToFirst")
            is equivalent to FindElement({Name:"Something"}, Acc.TreeScope.SubTree,, Acc.TreeTraversalOptions.LastToFirst)
            is equivalent to FindElement({Name:"Something"}, 7,, 1)

        {Name:"Something"} => Name must match "Something" (case sensitive)
        {Name:"Something", matchmode:"SubString", casesensitive:False} => Name must contain "Something" anywhere inside the Name, case insensitive. matchmode:"SubString" == matchmode:2 == matchmode:Acc.MatchMode.SubString
        {Name:"Something", RoleText:"something else"} => Name must match "Something" and RoleText must match "something else"
        [{Name:"Something", Role:42}, {Name:"Something2", RoleText:"something else"}] => Name=="Something" and Role==42 OR Name=="Something2" and RoleText=="something else"
        {Name:"Something", not:[{RoleText:"something", mm:"Substring"}, {RoleText:"something else", cs:1}]} => Name must match "something" and RoleText cannot match "something" (with matchmode=Substring == matchmode=2) nor "something else" (casesensitive matching)
        {or:[{Name:"Something"},{Name:"Something else"}], or2:[{Role:20},{Role:42}]}
        {Location:{w:200, h:100, r:"client"}} => Location must match width 200 and height 100 relative to client

    Dump(scope:=1, delimiter:=" ", depth:=-1)
        Outputs relevant information about the element (Name, Value, Location etc)
        Scope is the search scope: 1=element itself; 2=direct children; 4=descendants (including children of children); 7=whole subtree (including element)
            The scope is additive: 3=element itself and direct children.
    DumpAll(delimiter:=" ", depth:=-1)
        Outputs relevant information about the element and all descendants of the element. This is equivalent to Dump(7)
    Highlight(showTime:=unset, color:="Red", d:=2)
        Highlights the element for a chosen period of time
        Possible showTime values:
            Unset - highlights for 2 seconds, or removes the highlighting
            0 - Indefinite highlighting. If the element object gets destroyed, so does the highlighting.
            Positive integer (eg 2000) - will highlight and pause for the specified amount of time in ms
            Negative integer - will highlight for the specified amount of time in ms, but script execution will continue
            "clear" - removes the highlight unconditionally
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
    Navigate(navDir)
        Navigates in one of the directions specified by Acc.NavDir constants. Not all elements implement this method.
    HitTest(x, y)
        Retrieves the child element or child object that is displayed at a specific point on the screen.
        This shouldn't be used, since Acc.ObjectFromPoint uses this internally
     
