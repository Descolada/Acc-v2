# Acc-v2
Acc library for AHK v2

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