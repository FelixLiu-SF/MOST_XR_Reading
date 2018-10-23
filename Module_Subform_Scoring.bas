Attribute VB_Name = "Module_Subform_Scoring"
Option Compare Database
Option Explicit

'---COUNTVISIBLE---'
Public Function CountVisible(FormName As String, SubFormControlName As String, ControlArray() As String) As Integer

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
        If Forms(FormName).Controls(SubFormControlName).Form.Controls(ControlArray(Index)).Visible Then
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
Public Function CountUnlocked(FormName As String, SubFormControlName As String, ControlArray() As String) As Integer

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
        If Not Forms(FormName).Controls(SubFormControlName).Form.Controls(ControlArray(Index)).Locked Then
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
Public Function CountScored(FormName As String, SubFormControlName As String, ControlArray() As String) As Integer

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
        If Not Forms(FormName).Controls(SubFormControlName).Form.Controls(ControlArray(Index)).Locked Then

            'Count scored (on unlocked boxes only)
            If Len(Nz(Forms(FormName).Controls(SubFormControlName).Form.Controls(ControlArray(Index)).Value,"")) > 0 Then
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
Public Function CountUnscored(FormName As String, SubFormControlName As String, ControlArray() As String) As Integer

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
        If Not Forms(FormName).Controls(SubFormControlName).Form.Controls(ControlArray(Index)).Locked Then

            'Count scored (on unlocked boxes only)
            If Len(Nz(Forms(FormName).Controls(SubFormControlName).Form.Controls(ControlArray(Index)).Value,"")) < 1 Then
                nUnscored = nUnscored + 1
            End If

        End If
    Next

    CountUnscored = nUnscored

    Exit Function

    ErrorHandler1:
        On Error Goto 0
        Exit Function

End Function

'---INSERTSCORE---'
Public Function InsertScore(FormName As String, SubFormControlName As String, ControlName As String, VariableName As String, TableName As String, FilterName As String, FilterValue As String)

  Dim SQLText as String
  Dim ScoreValue As String
  Dim SQLValue As String

  On Error GoTo ScoreErr

  Set db = DBEngine(0)(0)

  'Get the score value
  ScoreValue = Nz(Forms(FormName).Controls(SubFormControlName).Form.Controls(ControlName).Value,"")
  If Len(ScoreValue) < 1 Then
    SQLValue = "NULL"
  Else
    SQLValue = ScoreValue
  End If
  SQLValue = Trim(SQLValue)

  'Construct SQL code for insert updated score value
  SQLText = "UPDATE " & TableName & " SET " & TableName & "." & VariableName & " = " & SQLValue & " WHERE ((" & TableName & "." & FilterName & ")=" & FilterValue & ");"

  'Execute SQL update code
  DoCmd.SetWarnings False
  db.Execute SQLText
  DoCmd.SetWarnings True

  DirtySave(FormName)

  On Error GoTo 0
  Exit Function

ScoreErr:
  Resume Next

End Function

'---INSERTSCORE2---'
Public Function InsertScore2(FormName As String, SubFormControlName As String, ControlName As String, VariableName As String, TableName As String, FilterName1 As String, FilterValue1 As String, FilterName2 As String, FilterValue2 As String)

  Dim SQLText as String
  Dim ScoreValue As String
  Dim SQLValue As String

  On Error GoTo ScoreErr

  Set db = DBEngine(0)(0)

  'Get the score value
  ScoreValue = Nz(Forms(FormName).Controls(SubFormControlName).Form.Controls(ControlName).Value,"")
  If Len(ScoreValue) < 1 Then
    SQLValue = "NULL"
  Else
    SQLValue = ScoreValue
  End If
  SQLValue = Trim(SQLValue)

  'Construct SQL code for insert updated score value
  SQLText = "UPDATE " & TableName & " SET " & TableName & "." & VariableName & " = " & SQLValue  & " WHERE (((" & TableName & "." & FilterName1 & ")=""" & FilterValue1 & """) AND ((" & TableName & "." & FilterName2 & ")=""" & FilterValue2 & """));"

  'Execute SQL update code
  DoCmd.SetWarnings False
  db.Execute SQLText
  DoCmd.SetWarnings True

  DirtySave(FormName)

  On Error GoTo 0
  Exit Function

ScoreErr:
  Resume Next

End Function

'---MAKE_CONTROLAFTERUPDATE_FUNC---'
Public Function Make_ControlAfterUpdate_Func(FormName As String, SubFormControlName As String, ControlName As String, VariableName As String, TableName As String, FilterName1 As String, FilterValue1 As String, FilterName2 As String, FilterValue2 As String) As String
'Concatente string for updating dropdown menu SQL query

    Dim AfterFuncStr As String

    AfterFuncStr = "=InsertScore2(""" & FormName & """,""" & SubFormControlName & """,""" & ControlName & """,""" & VariableName & """,""" & TableName & """,""" & FilterName1 & """,""" & FilterValue1 & """,""" & FilterName2 & """,""" & FilterValue2 & """)"

    Make_ControlAfterUpdate_Func = AfterFuncStr

End Function

'---MYLOOKUP2---'
Public Function MyLookup2(TableName As String, VariableName As String, FilterName1 As String, FilterValue1 As String, FilterName2 As String, FilterValue2 As String) As String
'Look up value from Access table using 2 string filter criteria
  Dim TableValue As String
  Dim VariableIn As String
  Dim FilterIn As String

  On Error GoTo LookupErr

  'Construct strings for DLookup query
  VariableIn = "[" & VariableName & "]"
  FilterIn = "[" & FilterName1 & "] = """ & FilterValue1 & """ AND [" & FilterName2 & "] = """ & FilterValue2 & """"

  'Query for value and return it
  TableValue = DLookup(VariableIn, TableName, FilterIn)

  MyLookup2 = TableValue

  On Error GoTo 0
  Exit Function

LookupErr:
  Resume Next

End Function

'---SETCOMBOVALUE---'
Public Function SetComboValue(FormName As String, SubFormControlName As String, ControlName As String, VariableName As String, TableName As String, FilterName1 As String, FilterValue1 As String, FilterName2 As String, FilterValue2 As String)

  Dim TableValue As String

  'Get value from table
  TableValue = Nz(MyLookup2(TableName, VariableName, FilterName1, FilterValue1, FilterName2, FilterValue2),"")

  'Update ComboBox value if value is not null
  If Len(TableValue)>0 Then
    Forms(FormName).Controls(SubFormControlName).Form.Controls(ControlName).Value = TableValue = TableValue
  End If

End Function
