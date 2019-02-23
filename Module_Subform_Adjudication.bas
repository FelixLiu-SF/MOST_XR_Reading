Attribute VB_Name = "Module_Subform_Adjudication"
Option Compare Database
Option Explicit

Global MOST_ADJ1_PARoot_Array(3) As String
Global MOST_ADJ2_PARoot_Array(16) As String
Global MOST_ADJ1_LATRoot_Array(4) As String
Global MOST_ADJ2_LATRoot_Array(14) As String
Global MOST_ADJ1_RV1234_XB_Vars() As String
Global MOST_ADJ2_RV1234_XB_Vars() As String
Global MOST_ADJ1_RV1234_LXB_Vars() As String
Global MOST_ADJ2_RV1234_LXB_Vars() As String

'---MOST_ADJ_LOAD_VARIABLENAMES---'
Public Function MOST_Adj_Load_VariableNames()
  'Call MOST_Load_VariableNameArrays from Module_MOST_Variables before calling this function

  'Define list of priority 1 and priority 2 adjudication variable name roots
  MOST_ADJ1_PARoot_Array(0) = "TFKLG"
  MOST_ADJ1_PARoot_Array(1) = "TFJSM"
  MOST_ADJ1_PARoot_Array(2) = "TFJSL"

  MOST_ADJ2_PARoot_Array(0) = "OSFM"
  MOST_ADJ2_PARoot_Array(1) = "OSFL"
  MOST_ADJ2_PARoot_Array(2) = "OSTM"
  MOST_ADJ2_PARoot_Array(3) = "OSTL"
  MOST_ADJ2_PARoot_Array(4) = "SCFM"
  MOST_ADJ2_PARoot_Array(5) = "SCFL"
  MOST_ADJ2_PARoot_Array(6) = "SCTM"
  MOST_ADJ2_PARoot_Array(7) = "SCTL"
  MOST_ADJ2_PARoot_Array(8) = "CYFM"
  MOST_ADJ2_PARoot_Array(9) = "CYFL"
  MOST_ADJ2_PARoot_Array(10) = "CYTM"
  MOST_ADJ2_PARoot_Array(11) = "CYTL"
  MOST_ADJ2_PARoot_Array(12) = "ATTM"
  MOST_ADJ2_PARoot_Array(13) = "ATTL"
  MOST_ADJ2_PARoot_Array(14) = "CHOM"
  MOST_ADJ2_PARoot_Array(15) = "CHOL"

  MOST_ADJ1_LATRoot_Array(0) = "PFKLG"
  MOST_ADJ1_LATRoot_Array(1) = "PFJSN"
  MOST_ADJ1_LATRoot_Array(2) = "FTJSM"
  MOST_ADJ1_LATRoot_Array(3) = "FTJSL"

  MOST_ADJ2_LATRoot_Array(0) = "OSFA"
  MOST_ADJ2_LATRoot_Array(1) = "OSFP"
  MOST_ADJ2_LATRoot_Array(2) = "OSPS"
  MOST_ADJ2_LATRoot_Array(3) = "OSPI"
  MOST_ADJ2_LATRoot_Array(4) = "OSTA"
  MOST_ADJ2_LATRoot_Array(5) = "OSTP"
  MOST_ADJ2_LATRoot_Array(6) = "SCPF"
  MOST_ADJ2_LATRoot_Array(7) = "CYPF"
  MOST_ADJ2_LATRoot_Array(8) = "CHON"
  MOST_ADJ2_LATRoot_Array(9) = "JE"
  MOST_ADJ2_LATRoot_Array(10) = "OSQI"
  MOST_ADJ2_LATRoot_Array(11) = "OPTU"
  MOST_ADJ2_LATRoot_Array(12) = "OPTL"
  MOST_ADJ2_LATRoot_Array(13) = "OSLB"

  'Create PA view adjudication variable list
  MOST_ADJ1_RV1234_XB_Vars = Concat_VisitVarSide(MOST_Visits_Array, MOST_PAKnee_Array, MOST_ADJ1_PARoot_Array)
  MOST_ADJ2_RV1234_XB_Vars = Concat_VisitVarSide(MOST_Visits_Array, MOST_PAKnee_Array, MOST_ADJ2_PARoot_Array)

  'Create LAT view adjudication variable list
  MOST_ADJ1_RV1234_LXB_Vars = Concat_VisitVarSide(MOST_Visits_Array, MOST_LATKnee_Array, MOST_ADJ1_LATRoot_Array)
  MOST_ADJ2_RV1234_LXB_Vars = Concat_VisitVarSide(MOST_Visits_Array, MOST_LATKnee_Array, MOST_ADJ2_LATRoot_Array)

End Function

'---ISADJUDICATION----'
Public Function IsAdjudication() As Boolean

  Dim AdjudicationFlag As Integer
  Dim AdjudicationBoolean As Boolean

  'Get DebugFlag
  AdjudicationFlag = DLookup("AdjudicationFlag", "tblProperties", "RecordID = 1")
  If Len(Nz(AdjudicationFlag, "")) > 0 Then
      If AdjudicationFlag > 0 Then
          AdjudicationBoolean = True
      Else
          AdjudicationBoolean = False
      End If

  Else
      AdjudicationBoolean = False
  End If

  IsAdjudication = AdjudicationFlag

End Function

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
    Dim ColorLightBlue As Long

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
    ColorLightBlue = RGB(166,166,255)

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
              Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).BackColor = ColorLightBlue
              Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).FontWeight = 400
              Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).BorderColor = ColorSilver

            Else
              'Adjudication character not found
              Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).BackColor = ColorWhite
              Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).FontWeight = 400
              Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).BorderColor = ColorSilver
            End If

        Else 'Is Null

            Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).BackColor = ColorLightBlue
            Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).FontWeight = 400
            Forms(FormName).Controls(SubFormControlName).Controls(ControlIn).BorderColor = ColorSilver

        End If 'null

    End If 'locked

    AdjBackcolorCodePriority1 = True

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
    Dim ColorLightBlue As Long

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
    ColorLightBlue = RGB(166,166,255)

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

    AdjBackcolorCodePriority2 = True

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

'---SETADJ1COMBOUPDATE_RV1234----'
Public Function SetAdj1ComboUpdate_RV1234(FormName As String, SubFormControlName As String, ViewPrefix As String, VarNameRoot As String, TableName As String, FilterName1 As String, FilterValue1 As String)

    Dim DummyBoolean As Boolean
    Dim VisitArray(4) As String
    Dim VisitNum(4) As Integer
    Dim ControlName As String
    Dim VariableName As String
    Dim FilterName2 As String
    Dim FilterValue2 As String
    Dim AfterUpdateStr As String
    Dim Index As Integer

    'Define default variables
    VisitArray(0) = "RV1"
    VisitArray(1) = "RV2"
    VisitArray(2) = "RV3"
    VisitArray(3) = "RV4"

    VisitNum(0) = 1
    VisitNum(1) = 2
    VisitNum(2) = 3
    VisitNum(3) = 4

    'Loop through visits
    Index = 0
    For Index = 0 To 4

        'Construct Variable name
        VariableName = ViewPrefix & VarNameRoot

        'Construct ComboBox Control name
        ControlName = "Combo_" & VisitArray(Index) & ViewPrefix & VarNameRoot

        'Construct visit filters
        FilterName2 = "RVNUM"
        FilterValue2 = CStr(VisitNum(Index))

        'Construct ComboBox selection string
        AfterUpdateStr = Make_Adj1ControlAfterUpdate_Func(FormName, SubFormControlName, ControlName, VariableName, TableName, FilterName1, FilterValue1, FilterName2, FilterValue2)

        'Set the after update string to the OnFocus property of the ComboBox
        DummyBoolean = Control_Edit_AfterUpdate(FormName, SubFormControlName, ControlName, AfterUpdateStr)

    Next

End Function

'---SETADJ2COMBOUPDATE_RV1234----'
Public Function SetAdj2ComboUpdate_RV1234(FormName As String, SubFormControlName As String, ViewPrefix As String, VarNameRoot As String, TableName As String, FilterName1 As String, FilterValue1 As String)

    Dim DummyBoolean As Boolean
    Dim VisitArray(4) As String
    Dim VisitNum(4) As Integer
    Dim ControlName As String
    Dim VariableName As String
    Dim FilterName2 As String
    Dim FilterValue2 As String
    Dim AfterUpdateStr As String
    Dim Index As Integer

    'Define default variables
    VisitArray(0) = "RV1"
    VisitArray(1) = "RV2"
    VisitArray(2) = "RV3"
    VisitArray(3) = "RV4"

    VisitNum(0) = 1
    VisitNum(1) = 2
    VisitNum(2) = 3
    VisitNum(3) = 4

    'Loop through visits
    Index = 0
    For Index = 0 To 4

        'Construct Variable name
        VariableName = ViewPrefix & VarNameRoot

        'Construct ComboBox Control name
        ControlName = "Combo_" & VisitArray(Index) & ViewPrefix & VarNameRoot

        'Construct visit filters
        FilterName2 = "RVNUM"
        FilterValue2 = CStr(VisitNum(Index))

        'Construct ComboBox selection string
        AfterUpdateStr = Make_Adj2ControlAfterUpdate_Func(FormName, SubFormControlName, ControlName, VariableName, TableName, FilterName1, FilterValue1, FilterName2, FilterValue2)

        'Set the after update string to the OnFocus property of the ComboBox
        DummyBoolean = Control_Edit_AfterUpdate(FormName, SubFormControlName, ControlName, AfterUpdateStr)

    Next

End Function
