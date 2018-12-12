Attribute VB_Name = "Module_Buttons"
Option Compare Database
Option Explicit

'---BUTTONNEXT---'
Public Function ButtonNext(FormName As String, SignVarName As String)

    Dim SignCheck As Variant
    Dim nMaxRec As Integer
    Dim Debug_Flag As Integer
    Dim MsgResponse As Integer

    On Error Goto ErrorHandler1

    'Preallocate maximum index to 1
    nMaxRec = 1

    'Get properties
    Debug_Flag = DLookup("DebugFlag","tblProperties","RecordID = 1")
    nMaxRec = DLookup("MaxRecord","tblProperties","RecordID = 1")

    'Set Focus on Form
    Forms(FormName).SetFocus

    If Debug_Flag < 1 Then
    'no debug - proceed as normal

        'check if current record is signed
        SignCheck = Forms(FormName).Recordset.Fields(SignVarName).Value

        If Forms(FormName).CurrentRecord < nMaxRec And Len(Nz(SignCheck, "")) > 0 Then
            'signed and not max record - continue to next record
            DoCmd.GoToRecord , , acNext

        ElseIf Len(Nz(SignCheck,"")) < 1 Then
            'not signed - ask user for confirmation
            MsgResponse = MsgBox("Current record is not signed. Are you sure you want to switch records?", vbYesNo + vbCritical + vbDefaultButton2, "Quit")
            If Forms(FormName).CurrentRecord < nMaxRec And MsgResponse = vbYes Then
                'answer is yes and not at max index - go to next
                DoCmd.GoToRecord , , acNext

            Else
                'do not go to next record
                Exit Function
            End If
        Else
            ' do not go to next record
            Exit Function
        End If

    Else
    ' debug mode, just go to next record if not at max index

        If Forms(FormName).CurrentRecord < nMaxRec Then
            DoCmd.GoToRecord , , acNext
        Else
            Exit Function
        End If

    End If

    On Error GoTo 0
    Exit Function

    ErrorHandler1:
      Exit Function

End Function

'---BUTTONPREV---'
Public Function ButtonPrev(FormName As String, SignVarName As String)

    Dim SignCheck As Variant
    Dim Debug_Flag As Integer
    Dim MsgResponse As Integer

    On Error Goto ErrorHandler1

    'Get properties
    Debug_Flag = DLookup("DebugFlag","tblProperties","RecordID = 1")

    'Set Focus on Form
    Forms(FormName).SetFocus

    If Debug_Flag < 1 Then
    'no debug - proceed as normal

        'check if current record is signed
        SignCheck = Forms(FormName).Recordset.Fields(SignVarName).Value

        If Forms(FormName).CurrentRecord <> 1  And Len(Nz(SignCheck, "")) > 0 Then
            'signed and not first record - continue to previous record
            DoCmd.GoToRecord , , acPrevious

        ElseIf Len(Nz(SignCheck,"")) < 1 Then
            'not signed - ask user for confirmation
            MsgResponse = MsgBox("Current record is not signed. Are you sure you want to switch records?", vbYesNo + vbCritical + vbDefaultButton2, "Quit")
            If Forms(FormName).CurrentRecord <> 1 And MsgResponse = vbYes Then
                'answer is yes and not at max index - go to previous
                DoCmd.GoToRecord , , acPrevious

            Else
                'do not go to previous record
                Exit Function
            End If
        Else
            ' do not go to previous record
            Exit Function
        End If

    Else
    ' debug mode, just go to previous record if not at first index

        If Forms(FormName).CurrentRecord <> 1 Then
            DoCmd.GoToRecord , , acPrevious
        Else
            Exit Function
        End If

    End If

    On Error GoTo 0
    Exit Function

    ErrorHandler1:
      Exit Function

End Function

Public Function QuitRequest(FormName As String, SignVarName As String)

    Dim SignCheck As Variant
    Dim Debug_Flag As Integer
    Dim MsgResponse As Integer

    On Error Goto ErrorHandler1

    'Get properties
    Debug_Flag = DLookup("DebugFlag","tblProperties","RecordID = 1")

    'Set Focus on Form
    Forms(FormName).SetFocus

    If Debug_Flag < 1 Then
    'no debug - proceed as normal

        'check if current record is signed
        SignCheck = Forms(FormName).Recordset.Fields(SignVarName).Value

        If Len(Nz(SignCheck, "")) > 0 Then
            'signed record - continue with quitting

            'ADD PRE-QUIT CODE HERE
            QuitRequest = True

        ElseIf Len(Nz(SignCheck,"")) < 1 Then
            'not signed - ask user for confirmation
            MsgResponse = MsgBox("current record is not signed. Are you sure you want to quit?", vbYesNo + vbCritical + vbDefaultButton2, "Quit")
            If MsgResponse = vbYes Then
                'answer is yes - continue with quitting

                'ADD PRE-QUIT CODE HERE
                QuitRequest = True

            Else
                'do not quit
                QuitRequest = False
                Exit Function
            End If
        Else
            'do not quit
            QuitRequest = False
            Exit Function
        End If

    Else
    ' debug mode, continue with quitting

        'ADD PRE-QUIT CODE HERE
        QuitRequest = True

    End If

    On Error GoTo 0
    Exit Function

    ErrorHandler1:
      Exit Function

End Function
