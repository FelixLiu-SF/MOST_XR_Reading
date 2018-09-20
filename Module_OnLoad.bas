Attribute VB_Name = "Module_OnLoad"
Option Compare Database
Option Explicit

Public Function LoadForm(FormName As String)

    Dim Debug_Flag As Integer

    On Error GoTo ErrorHandler1

    Debug_Flag = DLookup("DebugFlag","tblDebug","RecordID = 1")

    If Debug_Flag < 1 Then

        'Set Focus on Form
        Forms(FormName).SetFocus

        'Maximize the form window
        DoCmd.Maximize

        'Hide the Microsoft Ribbon
        DoCmd.ShowToolbar "Ribbon", acToolbarNo

    End If

    'clear error object
    On Error GoTo -1

    On Error Goto ErrorHandler2

    'Load the DAO objects
    LoadDAO

    'Load EFilm Automation object
    LoadEfilmAuto

    On Error GoTo 0
    Exit Function

    ErrorHandler1:

        'clear error object
        On Error GoTo -1

        On Error Goto ErrorHandler2

        'Maximize the form window
        DoCmd.Maximize

        'Hide the Microsoft Ribbon
        DoCmd.ShowToolbar "Ribbon", acToolbarNo

        Resume Next

    ErrorHandler2:

        'Load the DAO objects
        LoadDAO

        'Load EFilm Automation object
        LoadEfilmAuto

        Exit Function

End Function


'---LOADRECNAV---'
Public Function LoadRecNav(FormName As String, TextBoxName As String)
'Refresh text boxes to show current record number

    Dim nCurRec As Integer

    On Error GoTo ErrorHandler1

      'Set Focus on Form
      Forms(FormName).SetFocus

      'Get current record number and update text box
      nCurRec = Forms(FormName).CurrentRecord
      Forms(FormName).Controls(TextBoxName).Value = CStr(nCurRec)

      Exit Function

    ErrorHandler1:
    Exit Function

End Function

'---LOADMAXREC---'
Public Function LoadMaxRec(FormName As String, TextBoxName As String)
'Refresh text boxes to show max record number

    Dim nMaxRec As Integer

    On Error GoTo ErrorHandler1

      'Set Focus on Form
      Forms(FormName).SetFocus

      'Get max record number and update text box
      nMaxRec = DLookup("MaxRecord","tblProperties","RecordID = 1")
      Forms(FormName).Controls(TextBoxName).Value = CStr(nMaxRec)

      Exit Function

    ErrorHandler1:
    Exit Function

End Function
