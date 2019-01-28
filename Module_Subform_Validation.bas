Attribute VB_Name = "Module_Subform_Validation"
Option Compare Database
Option Explicit

Global MOST_Validation_PA_JSN_Array(2) As String
Global MOST_Validation_PA_OST_Array(4) As String
Global MOST_Validation_PA_Other_Array(10) As String

Public Function MOST_Validation_VariableNameArrays()

    MOST_Validation_PA_JSN_Array(0) = "TFJSM"
    MOST_Validation_PA_JSN_Array(1) = "TFJSL"

    MOST_Validation_PA_OST_Array(0) = "OSFM"
    MOST_Validation_PA_OST_Array(1) = "OSFL"
    MOST_Validation_PA_OST_Array(2) = "OSTM"
    MOST_Validation_PA_OST_Array(3) = "OSTL"

    MOST_Validation_PA_Other_Array(0) = "SCFM"
    MOST_Validation_PA_Other_Array(1) = "SCFL"
    MOST_Validation_PA_Other_Array(2) = "SCTM"
    MOST_Validation_PA_Other_Array(3) = "SCTL"
    MOST_Validation_PA_Other_Array(4) = "CYFM"
    MOST_Validation_PA_Other_Array(5) = "CYFL"
    MOST_Validation_PA_Other_Array(6) = "CYTM"
    MOST_Validation_PA_Other_Array(7) = "CYTL"
    MOST_Validation_PA_Other_Array(8) = "ATTM"
    MOST_Validation_PA_Other_Array(9) = "ATTL"

End Function

Public Function IsMOSTNewCohortID(ReadingIDIn As String)

  Dim SiteStr As String
  Dim CohortChar As String
  Dim CohortNum As Integer
  Dim NewCohortFlag As Boolean

  NewCohortFlag = False

  'Check for clinic site indicator'
  SiteStr = Left(ReadingIDIn,2)

  If Nz(SiteStr,"") = "MB" Or Nz(SiteStr,"") = "MI" Then

    'Get the cohort digit indicator'
    CohortChar = Mid(ReadingIDIn,4,1)

    'Convert indicator to an integer'
    CohortNum = CInt(CohortChar)

    'Check if indicator is for MOST new cohort'
    If CohortNum >= 3 Then
      NewCohortFlag = True
    Else
      NewCohortFlag = False
    End If
  Else
    NewCohortFlag = False
  End If

  'Output boolean result'
  IsMOSTNewCohortID = NewCohortFlag

End Function

Public Function MOST_Validate_Features_Max(ReadingIDIn As String, VisitStrIn As String, SideView As String, VarArrayIn As String) As Variant

  Dim Index As Integer
  Dim Upper As Integer
  Dim VarName As String
  Dim DummyBoolean As Boolean
  Dim DummyValueStr As String
  Dim DummyValueInt As Integer
  Dim ValueOut As Variant
  Dim TableName As String
  Dim FilterName1 As String
  Dim FilterName2 As String
  Dim CounterFlag As Boolean

  On Error GoTo ErrorHandler_Main1

  'Preset variables
  TableName = "tblScores"
  FilterName1 = "READINGID"
  FilterName2 = "RVNUM"
  CounterFlag = True

  'Loop through variables and save max value
  Index = 0
  Upper = UBound(VarArrayIn, 1) - 1

  For Index = 0 To Upper
    VarName = SideView & VarArrayIn(Index)
    DummyValueStr = MyLookup2(TableName, VarName, FilterName1, ReadingIDIn, FilterName2, VisitStrIn)
    If Nz(DummyValueStr,"") <> ""

      DummyValueInt = CInt(DummyValueStr)

      If CounterFlag Then
          If DummyValueInt >= -0.5 Then
            ValueOut = DummyValueInt
            CounterFlag = False
          End If
      Else
        DummyValueInt = CInt(DummyValueStr)
        If DummyValueInt > ValueOut
          ValueOut = DummyValueInt
        End If
      End If

    End If
  Next

  MOST_Validate_Features_Max = Nz(ValueOut,"")

  On Error GoTo 0
  Exit Sub

ErrorHandler_Main1:

  MOST_Validate_Features_Max = ""
  Exit Sub

End Function

Public Function MOST_Validate_Features_Min(ReadingIDIn As String, VisitStrIn As String, SideView As String, VarArrayIn As String) As Variant

  Dim Index As Integer
  Dim Upper As Integer
  Dim VarName As String
  Dim DummyBoolean As Boolean
  Dim DummyValueStr As String
  Dim DummyValueInt As Integer
  Dim ValueOut As Variant
  Dim TableName As String
  Dim FilterName1 As String
  Dim FilterName2 As String
  Dim CounterFlag As Boolean

  On Error GoTo ErrorHandler_Main1

  'Preset variables
  TableName = "tblScores"
  FilterName1 = "READINGID"
  FilterName2 = "RVNUM"
  CounterFlag = True

  'Loop through variables and save max value
  Index = 0
  Upper = UBound(VarArrayIn, 1) - 1

  For Index = 0 To Upper
    VarName = SideView & VarArrayIn(Index)
    DummyValueStr = MyLookup2(TableName, VarName, FilterName1, ReadingIDIn, FilterName2, VisitStrIn)
    If Nz(DummyValueStr,"") <> ""

      DummyValueInt = CInt(DummyValueStr)

      If CounterFlag Then
          If DummyValueInt >= -0.5 Then
            ValueOut = DummyValueInt
            CounterFlag = False
          End If
      Else
        DummyValueInt = CInt(DummyValueStr)
        If DummyValueInt < ValueOut
          ValueOut = DummyValueInt
        End If
      End If

    End If
  Next

  MOST_Validate_Features_Max = Nz(ValueOut,"")

  On Error GoTo 0
  Exit Sub

ErrorHandler_Main1:

  MOST_Validate_Features_Max = ""
  Exit Sub

End Function

Public Function MOST_Validate_By_ID(ReadingIDIn As String)

  Dim ValidationResult As New Collection

  'Check if existing or new cohort

  'Validate right knee PA

  'Validate left knee PA

  'Validate right knee lateral

  'Validate left knee lateral

  'Return results

End Function

Public Function MOST_Validate_PA_Standard(ReadingIDIn As String, VisitStrIn As String, SideView As String, ByRef ValidationResult As Collection)

  Dim TableName As String
  Dim KLGVarName As String
  Dim FilterName1 As String
  Dim FilterName2 As String
  Dim KLGValue As String

  Dim FormName As String
  Dim SubFormControlName As String
  Dim KLGCombo As String
  Dim ComboVisible As Boolean
  Dim ComboUnlocked As Boolean

  Dim ValidationResultInt As Integer
  Dim ValidationResultStr As String
  Dim ValidationItemInt As String
  Dim ValidationItemStr As String

  Dim JSNMax As Variant
  Dim JSNMin As Variant
  Dim OstMax As Variant
  Dim OstMin As Variant
  Dim OthMax As Variant
  Dim OthMin As Variant

  'Preset variables
  TableName = "tblScores"
  KLGVarName = SideView & "TFKLG"
  FilterName1 = "READINGID"
  FilterName2 = "RVNUM"

  FormName = "Form_MOST_144_168"
  SubFormControlName = "Subform_PA"

  ValidationItemInt = VisitStrIn & SideView & "Int"
  ValidationItemStr = VisitStrIn & SideView & "Str"

  'Initialize validation result variables
  ValidationResultInt = 0
  ValidationResultStr = ""

  'Check if combobox is visible & unlocked
  KLGCombo = "Combo_" & VisitStrIn & KLGVarName

  ComboVisible = Forms(FormName).Controls(SubFormControlName).Form.Controls(KLGCombo).Visible
  ComboUnlocked = Not Forms(FormName).Controls(SubFormControlName).Form.Controls(KLGCombo).Locked

  'Continue if combo box is unlocked and visible
  If ComboVisible And ComboUnlocked

    'Get KLG value
    KLGValue = MyLookup2(TableName, KLGVarName, FilterName1, ReadingIDIn, FilterName2, VisitStrIn)

    'Check if empty
    If Nz(KLGValue,"") <> "" Then

      'Check if special missing value
      If KLGValue <> "-6" And KLGValue <> "-7" And KLGValue <> "-8" And KLGValue <> "-9" Then

        'Calculate feature mins and maxes
        JSNMax = MOST_Validate_Features_Max(ReadingIDIn, VisitStrIn, SideView, MOST_Validation_PA_JSN_Array)
        JSNMin = MOST_Validate_Features_Min(ReadingIDIn, VisitStrIn, SideView, MOST_Validation_PA_JSN_Array)
        OSTMax = MOST_Validate_Features_Max(ReadingIDIn, VisitStrIn, SideView, MOST_Validation_PA_OST_Array)
        OSTMin = MOST_Validate_Features_Min(ReadingIDIn, VisitStrIn, SideView, MOST_Validation_PA_OST_Array)
        OthMax = MOST_Validate_Features_Max(ReadingIDIn, VisitStrIn, SideView, MOST_Validation_PA_Other_Array)
        OthMin = MOST_Validate_Features_Min(ReadingIDIn, VisitStrIn, SideView, MOST_Validation_PA_Other_Array)

        'Continue to triage based on KLG value if not special missing value
        Select Case KLGValue

          Case "0"

          Case "1"

          Case "1.9"

          Case "2"

          Case "3"

          Case "4"

        End Select

      Else 'Output indicator for special missing value

        ValidationResultInt = 2
        ValidationResultStr = ValidationResultStr & ""

      End If 'missing value

    Else 'Output indicator for empty value

      ValidationResultInt = 2
      ValidationResultStr = ValidationResultStr & ""

    End If 'empty

    'Return checks as item and keys in referenced collection object

  End If 'combo locked

End Function