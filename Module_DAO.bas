Attribute VB_Name = "Module_DAO"
Option Compare Database
Option Explicit

'Create object variable for DAO database and recordset

Public db As DAO.Database
Public rs As DAO.Recordset

'---LOADDAO---'
Public Function LoadDAO()

    Dim str_err As String
    Dim msg_err As Integer

    On Error GoTo ErrorHandler1

    ' Load the DAO objects
    Set db = Nothing
    Set rs = Nothing

    On Error GoTo 0
    Exit Function

ErrorHandler1:
    'str_err = ""
    'msg_err = MsgBox(str_err, vbOKOnly, "DAO")

Exit Function

End Function
