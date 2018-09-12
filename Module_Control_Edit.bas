Attribute VB_Name = "Module_Control_Edit"
Option Compare Database
Option Explicit

'---CONTROL_EDIT_FOCUS---'
Public Function Control_Edit_OnFocus(FormIn as Access.Form, ControlName As String, FocusFuncStr As String)
'Insert function into control box to be called on focus

    FormIn(ControlName).OnGotFocus = FocusFuncStr

End Function

'---CONTROL_EDIT_LOSTFOCUS---'
Public Function Control_Edit_LostFocus(FormIn as Access.Form, ControlName As String, LostFuncStr As String)

    FormIn(ControlName).OnLostFocus = LostFuncStr

End Function

'---CONTROL_EDIT_AFTERUPDATE---'
Public Function Control_Edit_AfterUpdate(FormIn as Access.Form, ControlName As String, UpdateFuncStr As String)
'Insert function into control box to be called after updating

    FormIn(ControlName).AfterUpdate = UpdateFuncStr

End function

'---CONTROL_EDIT_BINDING---'
Public Function Control_Edit_Binding(FormIn as Access.Form, ControlName As String, ColCount As Integer,ColBind As Integer, ColWidth As String, LimitBool As Boolean)
'Update the control binding properties

    FormIn(ControlName).ColumnCount = ColCount 'display this many columns
    FormIn(ControlName).BoundColumn = ColBind 'store the value from this column
    FormIn(ControlName).ColumnWidths = ColWidth 'text specifying display widths
    FormIn(ControlName).LimitToList = LimitBool 'only allow values from this table

End Function

'---MAKE_CONTROLUPDATE_FUNC---'
Public Function Make_ControlUpdate_Func(ControlName As String, SelectFunc As String) As String
'Concatente string for updating dropdown menu SQL query

    Dim FocusFuncStr As String

    FocusFuncStr = "=UpdateDropdown(""" & ControlName & """,""" & SelectFunc & """)"

    Make_ControlUpdate_Func = FocusFuncStr

End Function

'---UPDATEDROPDOWN---'
Public Function UpdateDropdown(FormIn As Access.Form, ControlIn As String, ControlSQL As String)
' Update Combo Box object table if not Locked

    'dummy variables for artificial CPU wait
    Dim DumLoop As Integer
    Dim DumBool As Boolean

    If FormIn(ControlIn).Locked = False Then

        ' initialize DAO objects
        On Error GoTo ErrorHandler1
        Set db = DBEngine(0)(0)
        Set rs = db.OpenRecordset(ControlSQL)

        ' Update Combo Box
        FormIn(ControlIn).RowSourceType = "Table/Query"
        Set FormIn(ControlIn).Recordset = rs

        ' dummy loop to wait before clearing rs
        For DumLoop = 0 To 500
            DumBool = True
        Next

        ' clear the DAO objects
        Set rs = Nothing
        Set db = Nothing

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
