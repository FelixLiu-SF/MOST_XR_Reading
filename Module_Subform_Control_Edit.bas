Attribute VB_Name = "Module_Subform_Control_Edit"
Option Compare Database
Option Explicit

'---CONTROL_EDIT_FOCUS---'
Public Function Control_Edit_OnFocus(FormName As String, SubFormControlName As String, ControlName As String, FocusFuncStr As String)
'Insert function into control box to be called on focus

    Forms(FormName).Controls(SubFormControlName).Form.Controls(ControlName).OnGotFocus = FocusFuncStr

End Function

'---CONTROL_EDIT_LOSTFOCUS---'
Public Function Control_Edit_LostFocus(FormName As String, SubFormControlName As String, ControlName As String, LostFuncStr As String)

    Forms(FormName).Controls(SubFormControlName).Form.Controls(ControlName).OnLostFocus = LostFuncStr

End Function

'---CONTROL_EDIT_AFTERUPDATE---'
Public Function Control_Edit_AfterUpdate(FormName As String, SubFormControlName As String, ControlName As String, UpdateFuncStr As String)
'Insert function into control box to be called after updating

    Forms(FormName).Controls(SubFormControlName).Form.Controls(ControlName).AfterUpdate = UpdateFuncStr

End function

'---CONTROL_EDIT_BINDING---'
Public Function Control_Edit_Binding(FormName As String, SubFormControlName As String, ControlName As String, ColCount As Integer,ColBind As Integer, ColWidth As String, ComboListWidth As Single, LimitBool As Boolean)
'Update the control binding properties

    Dim TwipListWidth As Integer

    TwipListWidth = Int(ComboListWidth*1400)

    Forms(FormName).Controls(SubFormControlName).Form.Controls(ControlName).ColumnCount = ColCount 'display this many columns
    Forms(FormName).Controls(SubFormControlName).Form.Controls(ControlName).BoundColumn = ColBind 'store the value from this column
    Forms(FormName).Controls(SubFormControlName).Form.Controls(ControlName).ColumnWidths = ColWidth 'text specifying column display widths
    Forms(FormName).Controls(SubFormControlName).Form.Controls(ControlName).ListWidth = TwipListWidth 'test specifying list width
    Forms(FormName).Controls(SubFormControlName).Form.Controls(ControlName).LimitToList = LimitBool 'only allow values from this table

End Function

'---MAKE_CONTROLCOLOR_FUNC---'
Public Function Make_ControlColor_Func(FormName As String, SubFormControlName As String, ControlName As String) As String
'Concatente string for updating dropdown menu SQL query

    Dim ColorFuncStr As String

    ColorFuncStr = "=BackcolorCode(""" & FormName & """,""" & SubFormControlName & """,""" & ControlName & """)"

    Make_ControlColor_Func = ColorFuncStr

End Function

'---MAKE_CONTROLSAVE_FUNC---'
Public Function Make_ControlSave_Func(FormName As String) As String
'Concatente string for updating dropdown menu SQL query

    Dim SaveFuncStr As String

    SaveFuncStr = "=DirtySave(""" & FormName & """)"

    Make_ControlSave_Func = SaveFuncStr

End Function

'---MAKE_CONTROLUPDATE_FUNC---'
Public Function Make_ControlUpdate_Func(FormName As String, SubFormControlName As String, ControlName As String, SelectFunc As String) As String
'Concatente string for updating dropdown menu SQL query

    Dim FocusFuncStr As String

    FocusFuncStr = "=UpdateDropdown(""" & FormName & """,""" & SubFormControlName & """,""" & ControlName & """,""" & SelectFunc & """)"

    Make_ControlUpdate_Func = FocusFuncStr

End Function

'---UPDATEDROPDOWN---'
Public Function UpdateDropdown(FormName As String, SubFormControlName As String, ControlName As String, ControlSQL As String, ColCount As Integer,ColBind As Integer, ColWidth As String, ComboListWidth As Single, LimitBool As Boolean)
' Update Combo Box object table if not Locked

    'dummy variables for artificial CPU wait
    Dim DumLoop As Integer
    Dim DumBool As Boolean

    If Forms(FormName).Controls(SubFormControlName).Form.Controls(ControlName).Locked = False Then

        ' initialize DAO objects
        On Error GoTo ErrorHandler1
        Set db = DBEngine(0)(0)
        Set rs = db.OpenRecordset(ControlSQL)

        ' Update Combo Box query
        Forms(FormName).Controls(SubFormControlName).Form.Controls(ControlName).RowSourceType = "Table/Query"
        Set Forms(FormName).Controls(SubFormControlName).Form.Controls(ControlName).Recordset = rs

        ' dummy loop to wait before clearing rs
        For DumLoop = 0 To 500
            DumBool = True
        Next

        ' clear the DAO objects
        Set rs = Nothing
        Set db = Nothing

        'Set Combo Box binding
        DumBool = Control_Edit_Binding(FormName,SubFormControlName,ControlName,ColCount,ColBind,ColWidth,ComboListWidth, LimitBool)

        UpdateDropdown = True

        ' disable error handling
        On Error GoTo 0

    End If
    Exit Function

    ErrorHandler1:
        ' reload DAO objects
        LoadDAO

    Exit Function

End Function

'---DIRTYSAVE---'
Public Function DirtySave(FormName As String)
'Save Dirty record

    On Error GoTo DirtySaveErr

    If Forms(FormName).Dirty Then
        Forms(FormName).Dirty = False
    End If

    DoCmd.Save acForm, Forms!FormName

    On Error GoTo 0
    DirtySave = True
    Exit Function

    DirtySaveErr:
    Resume Next

End Function
