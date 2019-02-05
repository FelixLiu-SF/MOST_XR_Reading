Attribute VB_Name = "Module_Subform_Adjudication"
Option Compare Database
Option Explicit

'---ADJBACKCOLORCODEPRIORITY1---'
Public Function AdjBackcolorCodePriority1(FormName As String, SubFormControlName As String, ControlIn As String)
'Set the control background color according to property status

    'Color variables
    Dim ColorWhite As Long
    Dim ColorSilver As Long
    Dim ColorLightGrey As Long
    Dim ColorPaleYellow As Long
    Dim ColorYellow As Long
    Dim ColorBlue As Long
    Dim ColorBrown As Long

    'Value variables
    Dim CheckVar As Variant
    Dim CheckValue As Variant
    Dim CheckType As Integer
    Dim StrValue As String

    Dim SearchChar As String
    Dim SearchResult As Variant

    'Assign RGB color codes
    ColorWhite = RGB(255, 255, 255)
    ColorSilver = RGB(192, 192, 192)
    ColorLightGrey = RGB(210, 210, 210)
    ColorPaleYellow = RGB(255, 255, 166)
    ColorYellow = RGB(192, 192, 0)
    ColorBlue = RGB(0, 0, 150)
    ColorBrown = RGB(205, 133, 63)

    'Assign search character
    SearchChar = "/"

    'Check if the control is Locked
    If Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).Locked Then
    'Color the Locked=TRUE boxes

        Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).BackColor = ColorSilver
        Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).FontWeight = 400
        Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).BorderColor = ColorSilver

    Else 'Color the rest based on null values

        'Check for associated variable in control
        CheckVar = Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).ControlSource

        If Len(Nz(CheckVar, "")) > 0 Then
        'Control is bound to a variable

            'Get the bound variable type and value
            CheckType = VarType(CheckVar)
            CheckValue = Forms(FormName).Controls(SubFormControlName).Recordset.Fields(CheckVar).Value

            'Format value score into string
            If Len(Nz(CheckValue, "")) > 0 Then
                If CheckType = 2 Or CheckType = 3 Then
                    StrValue = Trim(CStr(CheckValue))
                Else
                    StrValue = Trim(CheckValue)
                End If
            Else
                StrValue = ""
            End If

        Else
        'Control is not bound to a variable. Assume value is a string

            StrValue = Nz(Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).Value, "")

        End If

        'Color the control background according to the value

        If Len(Nz(StrValue, "")) > 0 Then 'Is Not Null

            'Search value for adjudication character
            SearchResult = InStr(1,StrValue,SearchChar,1)

            If SearchResult > 0 Then
              'Adjudication character found, highlight the ComboBox
              Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).BackColor = ColorBlue
              Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).FontWeight = 400
              Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).BorderColor = ColorSilver

            Else
              'Adjudication character not found
              Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).BackColor = ColorWhite
              Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).FontWeight = 400
              Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).BorderColor = ColorSilver
            End If

        Else 'Is Null

            Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).BackColor = ColorBlue
            Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).FontWeight = 400
            Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).BorderColor = ColorSilver

        End If 'null

    End If 'locked

    BackcolorCode = True

End Function

'---ADJBACKCOLORCODEPRIORITY2---'
Public Function AdjBackcolorCodePriority2(FormName As String, SubFormControlName As String, ControlIn As String)
'Set the control background color according to property status

    'Color variables
    Dim ColorWhite As Long
    Dim ColorSilver As Long
    Dim ColorLightGrey As Long
    Dim ColorPaleYellow As Long
    Dim ColorYellow As Long
    Dim ColorBlue As Long
    Dim ColorBrown As Long

    'Value variables
    Dim CheckVar As Variant
    Dim CheckValue As Variant
    Dim CheckType As Integer
    Dim StrValue As String

    Dim SearchChar As String
    Dim SearchResult As Variant

    'Assign RGB color codes
    ColorWhite = RGB(255, 255, 255)
    ColorSilver = RGB(192, 192, 192)
    ColorLightGrey = RGB(210, 210, 210)
    ColorPaleYellow = RGB(255, 255, 166)
    ColorYellow = RGB(192, 192, 0)
    ColorBlue = RGB(0, 0, 150)
    ColorBrown = RGB(205, 133, 63)

    'Assign search character
    SearchChar = "/"

    'Check if the control is Locked
    If Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).Locked Then
    'Color the Locked=TRUE boxes

        Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).BackColor = ColorSilver
        Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).FontWeight = 400
        Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).BorderColor = ColorSilver

    Else 'Color the rest based on null values

        'Check for associated variable in control
        CheckVar = Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).ControlSource

        If Len(Nz(CheckVar, "")) > 0 Then
        'Control is bound to a variable

            'Get the bound variable type and value
            CheckType = VarType(CheckVar)
            CheckValue = Forms(FormName).Controls(SubFormControlName).Recordset.Fields(CheckVar).Value

            'Format value score into string
            If Len(Nz(CheckValue, "")) > 0 Then
                If CheckType = 2 Or CheckType = 3 Then
                    StrValue = Trim(CStr(CheckValue))
                Else
                    StrValue = Trim(CheckValue)
                End If
            Else
                StrValue = ""
            End If

        Else
        'Control is not bound to a variable. Assume value is a string

            StrValue = Nz(Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).Value, "")

        End If

        'Color the control background according to the value

        If Len(Nz(StrValue, "")) > 0 Then 'Is Not Null

            'Search value for adjudication character
            SearchResult = InStr(1,StrValue,SearchChar,1)

            If SearchResult > 0 Then
              'Adjudication character found, highlight the ComboBox
              Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).BackColor = ColorPaleYellow
              Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).FontWeight = 400
              Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).BorderColor = ColorSilver

            Else
              'Adjudication character not found
              Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).BackColor = ColorWhite
              Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).FontWeight = 400
              Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).BorderColor = ColorSilver
            End If

        Else 'Is Null

            Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).BackColor = ColorPaleYellow
            Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).FontWeight = 400
            Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).BorderColor = ColorSilver

        End If 'null

    End If 'locked

    BackcolorCode = True

End Function

'---MAKE_ADJ1CONTROLCOLOR_FUNC---'
Public Function Make_Adj1ControlColor_Func(FormName As String, SubFormControlName As String, ControlName As String) As String
'Concatente string for updating dropdown menu SQL query

    Dim ColorFuncStr As String

    ColorFuncStr = "=AdjBackcolorCodePriority1(""" & FormName & """,""" & SubFormControlName & """,""" & ControlName & """)"

    Make_Adj1ControlColor_Func = ColorFuncStr

End Function

'---MAKE_ADJ2CONTROLCOLOR_FUNC---'
Public Function Make_Adj2ControlColor_Func(FormName As String, SubFormControlName As String, ControlName As String) As String
'Concatente string for updating dropdown menu SQL query

    Dim ColorFuncStr As String

    ColorFuncStr = "=AdjBackcolorCodePriority2(""" & FormName & """,""" & SubFormControlName & """,""" & ControlName & """)"

    Make_Adj2ControlColor_Func = ColorFuncStr

End Function

'---MAKE_ADJ1CONTROLAFTERUPDATE_FUNC---'
Public Function Make_Adj1ControlAfterUpdate_Func(FormName As String, SubFormControlName As String, ControlName As String, VariableName As String, TableName As String, FilterName1 As String, FilterValue1 As String, FilterName2 As String, FilterValue2 As String) As String
'Concatente string for updating dropdown menu SQL query

    Dim AfterFuncStr As String

    AfterFuncStr = "=Adj1InsertScore2(""" & FormName & """,""" & SubFormControlName & """,""" & ControlName & """,""" & VariableName & """,""" & TableName & """,""" & FilterName1 & """,""" & FilterValue1 & """,""" & FilterName2 & """,""" & FilterValue2 & """)"

    Make_Adj1ControlAfterUpdate_Func = AfterFuncStr

End Function

'---MAKE_ADJ2CONTROLAFTERUPDATE_FUNC---'
Public Function Make_Adj2ControlAfterUpdate_Func(FormName As String, SubFormControlName As String, ControlName As String, VariableName As String, TableName As String, FilterName1 As String, FilterValue1 As String, FilterName2 As String, FilterValue2 As String) As String
'Concatente string for updating dropdown menu SQL query

    Dim AfterFuncStr As String

    AfterFuncStr = "=Adj2InsertScore2(""" & FormName & """,""" & SubFormControlName & """,""" & ControlName & """,""" & VariableName & """,""" & TableName & """,""" & FilterName1 & """,""" & FilterValue1 & """,""" & FilterName2 & """,""" & FilterValue2 & """)"

    Make_Adj2ControlAfterUpdate_Func = AfterFuncStr

End Function

'---ADJ1INSERTSCORE2---'
Public Function Adj1InsertScore2(FormName As String, SubFormControlName As String, ControlName As String, VariableName As String, TableName As String, FilterName1 As String, FilterValue1 As String, FilterName2 As String, FilterValue2 As String)

  Dim SQLText as String
  Dim ScoreValue As String
  Dim SQLValue As String
  Dim DummyBoolean As Boolean

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

  DummyBoolean = AdjBackcolorCodePriority1(FormName, SubFormControlName, ControlName)

  Set db = Nothing

  On Error GoTo 0
  Exit Function

ScoreErr:
  Resume Next

End Function

'---ADJ2INSERTSCORE2---'
Public Function Adj2InsertScore2(FormName As String, SubFormControlName As String, ControlName As String, VariableName As String, TableName As String, FilterName1 As String, FilterValue1 As String, FilterName2 As String, FilterValue2 As String)

  Dim SQLText as String
  Dim ScoreValue As String
  Dim SQLValue As String
  Dim DummyBoolean As Boolean

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

  DummyBoolean = AdjBackcolorCodePriority2(FormName, SubFormControlName, ControlName)

  Set db = Nothing

  On Error GoTo 0
  Exit Function

ScoreErr:
  Resume Next

End Function
