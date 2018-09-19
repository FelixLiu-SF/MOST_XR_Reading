Attribute VB_Name = "Module_OnLoad"
Option Compare Database
Option Explicit

Public Function LoadForm(FormName As String)

    Dim Debug_Flag As Integer

    On Error GoTo ErrorHandler1

    Debug_Flag = DLookup("DebugFlag","tblDebug","RecordID = 1")

    If Debug_Flag > 0 Then

        'Maximize the form window
        DoCmd.Maximize

        'Hide the Microsoft Ribbon
        DoCmd.ShowToolbar "Ribbon", acToolbarNo

    End If
    'Load the DAO objects
    LoadDAO

    'Load EFilm Automation object
    LoadEfilmAuto

    On Error GoTo 0
    Exit Function

    ErrorHandler1:

        'Maximize the form window
        DoCmd.Maximize

        'Hide the Microsoft Ribbon
        DoCmd.ShowToolbar "Ribbon", acToolbarNo

        Resume Next

End Function
