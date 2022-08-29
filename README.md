# Acc-v2
Acc library for AHK v2


## Notable changes (compared to Acc v1):
Acc v2 in not a port of AHK v1 Acc library, but instead a complete redesign to incorporate more object-oriented approaches. 

    1) All Acc elements are now array-like objects, where the "Length" property contains the number of children, 
    any nth children can be accessed with element[n], and children can be iterated over with for loops.
    2) Acc main functions are contained in the global Acc object
    3) Element methods are contained inside element objects
    4) Element properties can be used without the "acc" prefix
    5) ChildIds have been removed (are handled in the backend), but can be accessed through 
    el.ChildId
    6) Additional methods have been added for elements, such as FindFirst, FindAll, Click
    7) Acc constants are included in the Acc object
    8) AccViewer is built into the library: when ran directly the AccViewer will show, when included
    in another script then it won't show (but can be accessed by calling Acc.Viewer())

## Short introduction
Acc (otherwise known as IAccessible or MSAA) is a library to get information about (and sometimes interact with) windows, controls, and window elements that are otherwise not accessible with AHKs Control functions. 

When Acc.ahk is ran alone, it displays the AccViewer: a window inspecter to get Acc information from elements. In the AccViewer on the left side are displayed all the available properties for elements (Value, Name etc), and on the right side will be displayed the Acc tree which shows how elements are related to eachother. In the bottom is displayed the Acc path (right-click to copy it), which can be used to get that specific element with the Acc library.  

To get started with Acc, first [include](https://lexikos.github.io/v2/docs/commands/_Include.htm) it in your program with [#include](https://lexikos.github.io/v2/docs/commands/_Include.htm). If Acc.ahk is in the same folder as your script use `#include Acc.ahk`, if it is in the Lib folder then use `#include <Acc>`.  

All main Acc properties and methods are accessible through the global variable `Acc`, which is created when Acc.ahk is included in your script. 
To access Acc elements, first you need to get a starting point element (usually a window) and save it in a variable: something like `oAcc := Acc.ObjectFromWindow(WinTitle)`. WinTitle uses the same rules as any other AHK function, so SetTitleMatchMode also applies to it.  

To get elements from the window element, you can use the Acc path from AccViewer: for example `oEl := oAcc[4,1,4]` would get the windows 4th sub-element, then the sub-elements 1st child, and then its 4th child.  

All the properties in AccViewer can be accessed from the element: `oEl.Name` will get that elements' name or throw an error if the name doesn't exist. Most properties are read-only, but the Value property can sometimes be changed with `oEl.Value := "newvalue"`  

Element methods can be used in the same way. To do the default action (usually clicking), use `oEl.DoDefaultAction()`, to highlight the element for 2 seconds use `oEl.Highlight(2000)`. Dump info about a specific element with `oEl.Dump()` and dump info about sub-elements as well with `oEl.DumpAll()` (to get paths and info about all elements in a window, use it on the window element: `MsgBox( Acc.ObjectFromWindow(WinTitle).DumpAll() )`  

Some examples of how Acc.ahk can be used are included in the Examples folder.  

# Acc methods

    ObjectFromPoint(x:=unset, y:=unset, &idChild := "", activateChromium := True)
        Gets an Acc element from screen coordinates X and Y (NOT relative to the active window).
    ObjectFromWindow(hWnd:="A", idObject := 0, activateChromium := True)
        Gets an Acc element from a WinTitle, by default the active window. 
        Additionally idObject can be specified from Acc.OBJID constants (eg to get the Caret location).
    GetRootElement()
        Gets the Acc element for the Desktop
    ActivateChromiumAccessibility(hWnd) 
        Sends the WM_GETOBJECT message to the Chromium document element and waits for the 
        app to be accessible to Acc. This is called when ObjectFromPoint or ObjectFromWindow 
        activateChromium flag is set to True. A small performance increase may be gotten 
        if that flag is set to False when it is not needed.
    RegisterWinEvent(event, callback) 
        Registers an event (a constant from Acc.EVENT) to a callback function and returns a new object
            containing the WinEventHook
        The callback function needs to have three arguments: 
            CallbackFunction(oAcc, Event, EventTime)
        Unhooking of the event handler will happen once the WinEventHook object is destroyed
        (either when overwritten by a constant, or when the script closes).

    Legacy methods:
    SetWinEventHook(eventMin, eventMax, pCallback)
    UnhookWinEvent(hHook)
    ObjectFromPath(ChildPath, hWnd:="A")
        Same as ObjectFromWindow[comma-separated path]
    GetRoleText(nRole)
        Same as element.RoleText
    GetStateText(nState)
        Same as element.StateText
    Query(pAcc)
        For internal use
        
# Acc constants

    Constants can be accessed as properties (eg Acc.OBJID.CARET), or the property name can be
      accessed by getting as an item (eg Acc.OBJID[0xFFFFFFF8])
    OBJID
    STATE
    ROLE
    NAVDIR
    SELECTIONFLAG
    EVENT
Explanations for the constants are available [in Microsoft documentations](https://docs.microsoft.com/en-us/windows/win32/winauto/constants-and-enumerated-types).

# Acc element properties

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
    WinID               => ID of the window the element belongs to
    oAcc                => ComObject of the underlying IAccessible (this is usually not needed)
    childId             => childId of the underlying IAccessible (this is usually not needed)
    
# Acc element methods

    Select(flags)
        Modifies the selection or moves the keyboard focus of the specified object. flags can be any of the SELECTIONFLAG constants
    DoDefaultAction()
        Performs the specified object's default action. Not all objects have a default action.
    GetNthChild(n)
        This is equal to oElement[n]
    GetLocation(relativeTo:="")
        Returns an object containing the x, y coordinates and width and height: {x:x coordinate, y:y coordinate, w:width, h:height}. 
        relativeTo can be client, window or screen, default is A_CoordModeMouse.
    IsEqual(oCompare)
        Checks whether the element is equal to another element (oCompare)
    FindFirst(condition, scope:=4) 
        Finds the first element matching the condition (see description under ValidateCondition)
        Scope is the search scope: 1=element itself; 2=direct children; 4=descendants (including children of children)
        The returned element also has the "Path" property with the found elements path
    FindAll(condition, scope:=4)
        Returns an array of elements matching the condition (see description under ValidateCondition)
        The returned elements also have the "Path" property with the found elements path
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
    Dump()
        Outputs relevant information about the element (Name, Value, Location etc)
    DumpAll()
        Outputs relevant information about the element and all descendants of the element
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
    Navigate(navDir)
        Navigates in one of the directions specified by Acc.NAVDIR constants. Not all elements implement this method.
    HitTest(x, y)
        Retrieves the child element or child object that is displayed at a specific point on the screen.
        This shouldn't be used, since Acc.ObjectFromPoint uses this internally
     
