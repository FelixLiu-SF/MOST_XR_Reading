Attribute VB_Name = "Module_EFilm"
Option Compare Database
Option Explicit

'Create object variable for EFilm Automation Server

Public EFilmAuto As Object

'---LOADEFILMAUTO---'
Public Function LoadEFilmAuto()

    Dim str_err As String
    Dim msg_err As Integer

    On Error GoTo ErrorHandler1

    ' Load the EFilm Automation Server
    Set EFilmAuto = CreateObject("Efilm.Document")

    On Error GoTo 0
    Exit Function

ErrorHandler1:
    'str_err = ""
    'msg_err = MsgBox(str_err, vbOKOnly, "EFilm-Access Connection")

    Exit Function

End Function

'---UNLOADEFILMAUTO---'
Public Function UnloadEFilmAuto()

    On Error GoTo ErrorHandler1

    Set EFilmAuto = Nothing

    On Error Goto 0
    Exit Function

ErrorHandler1:
    'str_err = ""
    'msg_err = MsgBox(str_err, vbOKOnly, "EFilm-Access Unload")

Exit Function

End Function

'---VBA_OPENSTUDY---'
Public Function VBA_OpenStudy(PATID As String, PATACC As String, NUMXR As Integer) As Boolean

    Dim WindowBool As Boolean

    On Error GoTo ErrorHandler1

    WindowBool = EFilmAuto.oleShowMainWindow(1)
    VBA_OpenStudy = EFilmAuto.oleOpenStudy2(PATID, PATACC, True, False, 1, NUMXR, 1, 1, False, False, "{0CBB4846-0868-4f42-8AC3-63F5B8822AF6}")

    On Error GoTo 0
    Exit Function

ErrorHandler1:

    LoadEFilmAuto

Exit Function

End Function
