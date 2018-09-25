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

'---UNLOADRIBBON---'
Public Function UnloadRibbon(FormName As String)

    On Error GoTo ErrorHandler1

    'Set Focus on Form
    Forms(FormName).SetFocus

    'Show the Microsoft Ribbon
    DoCmd.ShowToolbar "Ribbon", acToolbarYes

    Exit Function

    ErrorHandler1:
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

'---SKIPSIGNED---'
Public Function SkipSigned(FormName As String, SignVarName As String)
'Skip until record without signing is reached (or EOF)

    Dim SignCheck As Variant
    Dim Index As Integer
    Dim nMaxRec As Integer
    Dim Debug_Flag As Integer

    On Error GoTo ErrorHandler1

    'Preallocate maximum index to 1
    nMaxRec = 1

    'Get properties
    Debug_Flag = DLookup("DebugFlag","tblDebug","RecordID = 1")

    If Debug_Flag < 1 Then

      'Set Focus on Form
      Forms(FormName).SetFocus

      'Check the first record
      DoCmd.GoToRecord , , acFirst
      SignCheck = Forms(FormName).Recordset.Fields(SignVarName).Value

      If Len(Nz(SignCheck,"")) > 0 Then
          'If first record is signed, move on to next records

          'Try to get the maximum record index AKA last record
          nMaxRec = DLookup("MaxRecord","tblProperties","RecordID = 1")
          DoCmd.GoToRecord , , acFirst
          SignCheck = Forms(FormName).Recordset.Fields(SignVarName).Value

          'Loop until a record is not signed
          For Index = 1 To nMaxRec
              If Forms(FormName).CurrentRecord < nMaxRec And Len(Nz(SignCheck, "")) > 0 Then

                  DoCmd.GoToRecord , , acNext
                  SignCheck = Forms(FormName).Recordset.Fields(SignVarName).Value

              ElseIf Forms(FormName).CurrentRecord = nMaxRec And Len(Nz(SignCheck, "")) > 0 Then

                  DoCmd.GoToRecord , , acFirst
                  SignCheck = Forms(FormName).Recordset.Fields(SignVarName).Value

              End If
          Next

      End If 'first record sign check

    End If 'debugflag

    On Error GoTo 0
    Exit Function

    ErrorHandler1:
        Exit Function

End Function
