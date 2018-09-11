Option Compare Database
Option Explicit

'---LOCKBOX---'
Public Function LockBox(FormIn as Access.Form, ControlIn As String, LockBool As Boolean)

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

If HideBool Then
    FormIn(ControlIn).Visible = False
Else
    FormIn(ControlIn).Visible = True
End If

LockBox = True

End function

'---HIDEVARS---'
Public Function HideVars(FormIn as Access.Form, ControlArray As String, HideBool As Boolean)

Dim Upper as Integer
Dim Index as Integer
Dim FuncBool as Boolean

Upper = UBound(ControlArray, 1)

Index = 0
For Index = 0 To Upper
    FuncBool = HideBox(FormIn,ControlArray(Index), HideBool)
Next

HideVars = True

End function

'---LOCKVARS---'
Public Function LockVars(FormIn as Access.Form, ControlArray As String, LockBool As Boolean)

Dim Upper as Integer
Dim Index as Integer
Dim FuncBool as Boolean

Upper = UBound(ControlArray, 1)

Index = 0
For Index = 0 To Upper
    FuncBool = LockBox(FormIn,ControlArray(Index), LockBool)
Next

LockVars = True

End function
