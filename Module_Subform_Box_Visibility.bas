Attribute VB_Name = "Module_Subform_Box_Visibility"
Option Compare Database
Option Explicit

'---LOCKBOX---'
Public Function LockBox(FormName As String, SubFormControlName As String, ControlName As String, LockBool As Boolean)
'Change the control properties for locked and enabled

    If LockBool Then
        Forms(FormName).Controls(SubFormControlName).Controls(ControlName).Enabled = False
        Forms(FormName).Controls(SubFormControlName).Controls(ControlName).Locked = True
    Else
        Forms(FormName).Controls(SubFormControlName).Controls(ControlName).Enabled = True
        Forms(FormName).Controls(SubFormControlName).Controls(ControlName).Locked = False
    End If

    LockBox = True

End function

'---HIDEBOX---'
Public Function HideBox(FormName As String, SubFormControlName As String, ControlName As String, HideBool As Boolean)
'Change the control properties for visibility

    If HideBool Then
        Forms(FormName).Controls(SubFormControlName).Controls(ControlName).Visible = False
    Else
        Forms(FormName).Controls(SubFormControlName).Controls(ControlName).Visible = True
    End If

    HideBox = True

End function

'---HIDEVARS---'
Public Function HideVars(FormName As String, SubFormControlName As String, ControlArray() As String, HideBool As Boolean)
'Loop through string array of control box names and hide

    Dim Upper as Integer
    Dim Index as Integer
    Dim FuncBool as Boolean

    'Get the upper bound of the control name string array
    Upper = UBound(ControlArray, 1) - 1

    'Loop and hide
    Index = 0
    For Index = 0 To Upper
        FuncBool = HideBox(FormName, SubFormControlName, CStr(ControlArray(Index)), HideBool)
    Next

    HideVars = True

End function

'---LOCKVARS---'
Public Function LockVars(FormName As String, SubFormControlName As String, ControlArray() As String, LockBool As Boolean)
'Loop through string array of control box names and lock

    Dim Upper as Integer
    Dim Index as Integer
    Dim FuncBool as Boolean

    'Get the upper bound of the control name string array
    Upper = UBound(ControlArray, 1) - 1

    'Loop and lock
    Index = 0
    For Index = 0 To Upper
        FuncBool = LockBox(FormName, SubFormControlName, CStr(ControlArray(Index)), LockBool)
    Next

    LockVars = True

End function

'---BACKCOLORCODE---'
Public Function BackcolorCode(FormName As String, SubFormControlName As String, ControlIn As String)
'Set the control background color according to property status

    'Color variables
    Dim ColorWhite As Long
    Dim ColorSilver As Long
    Dim ColorLightGrey As Long
    Dim ColorPaleYellow As Long
    Dim ColorYellow As Long
    Dim ColorBlue As Long
    Dim ColorBrown As Long

    'Value variables
    Dim CheckVar As Variant
    Dim CheckValue As Variant
    Dim CheckType As Integer
    Dim StrValue As String

    'Assign RGB color codes
    ColorWhite = RGB(255, 255, 255)
    ColorSilver = RGB(192, 192, 192)
    ColorLightGrey = RGB(210, 210, 210)
    ColorPaleYellow = RGB(255, 255, 166)
    ColorYellow = RGB(192, 192, 0)
    ColorBlue = RGB(0, 0, 150)
    ColorBrown = RGB(205, 133, 63)

    'Check if the control is Locked
    If Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).Locked Then
    'Color the Locked=TRUE boxes

        Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).BackColor = ColorSilver
        Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).FontWeight = 400
        Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).BorderColor = ColorSilver

    Else 'Color the rest based on null values

        'Check for associated variable in control
        CheckVar = Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).ControlSource

        If Len(Nz(CheckVar, "")) > 0 Then
        'Control is bound to a variable

            'Get the bound variable type and value
            CheckType = VarType(CheckVar)
            CheckValue = Forms(FormName).Controls(SubFormControlName).Recordset.Fields(CheckVar).Value

            'Format value score into string
            If Len(Nz(CheckValue, "")) > 0 Then
                If CheckType = 2 Or CheckType = 3 Then
                    StrValue = Trim(CStr(CheckValue))
                Else
                    StrValue = Trim(CheckValue)
                End If
            Else
                StrValue = ""
            End If

        Else
        'Control is not bound to a variable. Assume value is a string

            StrValue = Nz(Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).Text, "")

        End If

        'Color the control background according to the value

        If Len(Nz(StrValue, "")) > 0 Then 'Is Not Null

            Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).BackColor = ColorWhite
            Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).FontWeight = 400
            Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).BorderColor = ColorSilver

        Else 'Is Null

            Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).BackColor = ColorPaleYellow
            Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).FontWeight = 400
            Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).BorderColor = ColorSilver

        End If 'null

    End If 'locked

    BackcolorCode = True

End Function
