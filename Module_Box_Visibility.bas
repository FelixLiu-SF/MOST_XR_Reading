Attribute VB_Name = "Module_Box_Visibility"
Option Compare Database
Option Explicit

'---LOCKBOX---'
Public Function LockBox(FormIn as Access.Form, ControlIn As String, LockBool As Boolean)
'Change the control properties for locked and enabled

    If LockBool Then
        FormIn(ControlIn).Enabled = False
        FormIn(ControlIn).Locked = True
    Else
        FormIn(ControlIn).Enabled = True
        FormIn(ControlIn).Locked = False
    End If

    LockBox = True

End function

'---HIDEBOX---'
Public Function HideBox(FormIn as Access.Form, ControlIn As String, HideBool As Boolean)
'Change the control properties for visibility

    If HideBool Then
        FormIn(ControlIn).Visible = False
    Else
        FormIn(ControlIn).Visible = True
    End If

    HideBox = True

End function

'---HIDEVARS---'
Public Function HideVars(FormIn as Access.Form, ControlArray() As String, HideBool As Boolean)
'Loop through string array of control box names and hide

    Dim Upper as Integer
    Dim Index as Integer
    Dim FuncBool as Boolean

    'Get the upper bound of the control name string array
    Upper = UBound(ControlArray, 1) - 1

    'Loop and hide
    Index = 0
    For Index = 0 To Upper
        FuncBool = HideBox(FormIn,CStr(ControlArray(Index)), HideBool)
    Next

    HideVars = True

End function

'---LOCKVARS---'
Public Function LockVars(FormIn as Access.Form, ControlArray() As String, LockBool As Boolean)
'Loop through string array of control box names and lock

    Dim Upper as Integer
    Dim Index as Integer
    Dim FuncBool as Boolean

    'Get the upper bound of the control name string array
    Upper = UBound(ControlArray, 1) - 1

    'Loop and lock 
    Index = 0
    For Index = 0 To Upper
        FuncBool = LockBox(FormIn, CStr(ControlArray(Index)), LockBool)
    Next

    LockVars = True

End function
