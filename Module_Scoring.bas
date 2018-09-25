Attribute VB_Name = "Module_Scoring"
Option Compare Database
Option Explicit

'---COUNTVISIBLE---'
Public Function CountVisible(FormName As String, ControlArray() As String) As Integer

    Dim nVisible As Integer
    Dim Index As Integer
    Dim ControlUB As Integer

    On Error GoTo ErrorHandler1

    nVisible = 0
    Index = 0
    CountVisible = 0

    'Get upper bound for each string array
    ControlUB = UBound(ControlArray,1)

    'Loop through control array and check if visible
    For Index = 0 To (ControlUB - 1)
        If Forms(FormName).Controls(ControlArray(Index)).Visible Then
            nVisible = nVisible + 1
        End If
    Next

    CountVisible = nVisible

    Exit Function

    ErrorHandler1:
        On Error Goto 0
        Exit Function

End Function

'---COUNTUNLOCKED---'
Public Function CountUnlocked(FormName As String, ControlArray() As String) As Integer

    Dim nUnlocked As Integer
    Dim Index As Integer
    Dim ControlUB As Integer

    On Error GoTo ErrorHandler1

    nUnlocked = 0
    Index = 0
    CountUnlocked = 0

    'Get upper bound for each string array
    ControlUB = UBound(ControlArray,1)

    'Loop through control array and check if visible
    For Index = 0 To (ControlUB - 1)
        If Not Forms(FormName).Controls(ControlArray(Index)).Locked Then
            nUnlocked = nUnlocked + 1
        End If
    Next

    CountUnlocked = nUnlocked

    Exit Function

    ErrorHandler1:
        On Error Goto 0
        Exit Function

End Function

'---COUNTSCORED---'
Public Function CountScored(FormName As String, ControlArray() As String) As Integer

    Dim nScored As Integer
    Dim Index As Integer
    Dim ControlUB As Integer

    On Error GoTo ErrorHandler1

    nScored = 0
    Index = 0
    CountScored = 0

    'Get upper bound for each string array
    ControlUB = UBound(ControlArray,1)

    'Loop through control array and check if visible
    For Index = 0 To (ControlUB - 1)
        If Not Forms(FormName).Controls(ControlArray(Index)).Locked Then

            'Count scored (on unlocked boxes only)
            If Len(Nz(Forms(FormName).Controls(ControlArray(Index)).Value,"")) > 0 Then
                nScored = nScored + 1
            End If

        End If
    Next

    CountScored = nScored

    Exit Function

    ErrorHandler1:
        On Error Goto 0
        Exit Function

End Function

'---COUNTUNSCORED---'
Public Function CountUnscored(FormName As String, ControlArray() As String) As Integer

    Dim nUnscored As Integer
    Dim Index As Integer
    Dim ControlUB As Integer

    On Error GoTo ErrorHandler1

    nUnscored = 0
    Index = 0
    CountUnscored = 0

    'Get upper bound for each string array
    ControlUB = UBound(ControlArray,1)

    'Loop through control array and check if visible
    For Index = 0 To (ControlUB - 1)
        If Not Forms(FormName).Controls(ControlArray(Index)).Locked Then

            'Count scored (on unlocked boxes only)
            If Len(Nz(Forms(FormName).Controls(ControlArray(Index)).Value,"")) < 1 Then
                nUnscored = nUnscored + 1
            End If

        End If
    Next

    CountScored = nUnscored

    Exit Function

    ErrorHandler1:
        On Error Goto 0
        Exit Function

End Function
