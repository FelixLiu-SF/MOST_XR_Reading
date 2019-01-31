Attribute VB_Name = "Module_Subform_Validation"
Option Compare Database
Option Explicit

Global MOST_Validation_PA_JSN_Array(2) As String
Global MOST_Validation_PA_OST_Array(4) As String
Global MOST_Validation_PA_Other_Array(10) As String
Global MOST_Validation_LAT_JSN_Array(1) As String
Global MOST_Validation_LAT_OST_Array(3) As String
Global MOST_Validation_LAT_Other_Array(3) As String

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

    MOST_Validation_LAT_JSN_Array(0) = "PFJSN"

    MOST_Validation_LAT_OST_Array(0) = "OSFA"
    MOST_Validation_LAT_OST_Array(1) = "OSPS"
    MOST_Validation_LAT_OST_Array(2) = "OSPI"

    MOST_Validation_LAT_Other_Array(0) = "SCPF"
    MOST_Validation_LAT_Other_Array(1) = "CYPF"
    MOST_Validation_LAT_Other_Array(2) = "JE"

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

Public Function MOST_Validate_Features_Max(ReadingIDIn As String, VisitStrIn As String, SideView As String, VarArrayIn() As String) As Variant

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
    If Nz(DummyValueStr,"") <> "" Then

      DummyValueInt = CInt(DummyValueStr)

      If CounterFlag Then
          If DummyValueInt >= -0.5 Then
            ValueOut = DummyValueInt
            CounterFlag = False
          End If
      Else
        DummyValueInt = CInt(DummyValueStr)
        If DummyValueInt > ValueOut Then
          ValueOut = DummyValueInt
        End If
      End If

    End If
  Next

  MOST_Validate_Features_Max = Nz(ValueOut,"")

  On Error GoTo 0
  Exit Function

ErrorHandler_Main1:

  MOST_Validate_Features_Max = ""
  Exit Function

End Function

Public Function MOST_Validate_Features_Min(ReadingIDIn As String, VisitStrIn As String, SideView As String, VarArrayIn() As String) As Variant

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
    If Nz(DummyValueStr,"") <> "" Then

      DummyValueInt = CInt(DummyValueStr)

      If CounterFlag Then
          If DummyValueInt >= -0.5 Then
            ValueOut = DummyValueInt
            CounterFlag = False
          End If
      Else
        DummyValueInt = CInt(DummyValueStr)
        If DummyValueInt < ValueOut Then
          ValueOut = DummyValueInt
        End If
      End If

    End If
  Next

  MOST_Validate_Features_Min = Nz(ValueOut,"")

  On Error GoTo 0
  Exit Function

ErrorHandler_Main1:

  MOST_Validate_Features_Min = ""
  Exit Function

End Function

Public Function MOST_Validate_By_ID(ReadingIDIn As String) As Integer

  Dim ValidationResult As New Collection
  Dim NewCohortBoolean As Boolean
  Dim ValidationResponse As Integer
  Dim DummyBoolean As Boolean

  ValidationResponse = vbYes

  'Check if existing or new cohort
  NewCohortBoolean = IsMOSTNewCohortID(ReadingIDIn)

  If NewCohortBoolean = True Then
    'New cohort

    'Validate right knee PA
    DummyBoolean = MOST_Validate_PA_2N_Invalid(ReadingIDIn, "3", "XR", ValidationResult)
    DummyBoolean = MOST_Validate_PA_Standard(ReadingIDIn, "4", "XR", ValidationResult)

    'Validate left knee PA
    DummyBoolean = MOST_Validate_PA_2N_Invalid(ReadingIDIn, "3", "XL", ValidationResult)
    DummyBoolean = MOST_Validate_PA_Standard(ReadingIDIn, "4", "XL", ValidationResult)

    'Validate right knee lateral
    DummyBoolean = MOST_Validate_LAT_2N_Invalid(ReadingIDIn, "3", "LXR", ValidationResult)
    DummyBoolean = MOST_Validate_LAT_Standard(ReadingIDIn, "4", "LXR", ValidationResult)

    'Validate left knee lateral
    DummyBoolean = MOST_Validate_LAT_2N_Invalid(ReadingIDIn, "3", "LXL", ValidationResult)
    DummyBoolean = MOST_Validate_LAT_Standard(ReadingIDIn, "4", "LXL", ValidationResult)

  Else 'Existing cohort

    'Validate right knee PA
    DummyBoolean = MOST_Validate_PA_Standard(ReadingIDIn, "3", "XR", ValidationResult)
    DummyBoolean = MOST_Validate_PA_Standard(ReadingIDIn, "4", "XR", ValidationResult)

    'Validate left knee PA
    DummyBoolean = MOST_Validate_PA_Standard(ReadingIDIn, "3", "XL", ValidationResult)
    DummyBoolean = MOST_Validate_PA_Standard(ReadingIDIn, "4", "XL", ValidationResult)

    'Validate right knee lateral
    DummyBoolean = MOST_Validate_LAT_2N_Invalid(ReadingIDIn, "3", "LXR", ValidationResult)
    DummyBoolean = MOST_Validate_LAT_Standard(ReadingIDIn, "4", "LXR", ValidationResult)

    'Validate left knee lateral
    DummyBoolean = MOST_Validate_LAT_2N_Invalid(ReadingIDIn, "3", "LXL", ValidationResult)
    DummyBoolean = MOST_Validate_LAT_Standard(ReadingIDIn, "4", "LXL", ValidationResult)

  End If

  'Return results
  ValidationResponse = MOST_Validate_MsgBox(ValidationResult)

  MOST_Validate_By_ID = ValidationResponse

End Function

Public Function MOST_Validate_PA_Standard(ReadingIDIn As String, VisitStrIn As String, SideView As String, ByRef ValidationResult As Collection)

  Dim TableName As String
  Dim KLGVarName As String
  Dim FilterName1 As String
  Dim FilterName2 As String
  Dim KLGValue As String

  Dim TableName2 As String
  Dim VisitVarName As String
  Dim VisitName As String

  Dim FormName As String
  Dim SubFormControlName As String
  Dim KLGCombo As String
  Dim ComboVisible As Boolean
  Dim ComboUnlocked As Boolean

  Dim ValidationResultInt As Integer
  Dim ValidationResultStr As String
  Dim ValidationKeyInt As String
  Dim ValidationKeyStr As String

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

  TableName2 = "tblReadings"
  VisitVarName = "RV" & VisitStrIn & "TP"
  VisitName = MyLookup(TableName2, VisitVarName, FilterName1, ReadingIDIn)

  FormName = "Form_MOST_144_168"
  SubFormControlName = "Subform_PA"

  ValidationKeyInt = "RV" & VisitStrIn & SideView & "Int"
  ValidationKeyStr = "RV" & VisitStrIn & SideView & "Str"

  'Initialize validation result variables
  ValidationResultInt = 0
  ValidationResultStr = ""

  'Check if combobox is visible & unlocked
  KLGCombo = "Combo_" & "RV" & VisitStrIn & KLGVarName

  ComboVisible = Forms(FormName).Controls(SubFormControlName).Form.Controls(KLGCombo).Visible
  ComboUnlocked = Not Forms(FormName).Controls(SubFormControlName).Form.Controls(KLGCombo).Locked

  'Continue if combo box is unlocked and visible
  If ComboVisible And ComboUnlocked Then

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

        JSNMax = Nz(JSNMax,"")
        JSNMin = Nz(JSNMin,"")
        OSTMax = Nz(OSTMax,"")
        OSTMin = Nz(OSTMin,"")
        OthMax = Nz(OthMax,0)
        OthMin = Nz(OthMin,0)

        If JSNMax <> "" And JSNMin <> "" And OSTMax <> "" And OSTMin <> "" Then

          JSNMax = CInt(JSNMax)
          JSNMin = CInt(JSNMin)
          OSTMax = CInt(OSTMax)
          OSTMin = CInt(OSTMin)
          OthMax = CInt(OthMax)
          OthMin = CInt(OthMin)

          'Continue to triage based on KLG value if not special missing value
          Select Case KLGValue

            Case "0"

              If JSNMax > 0 Or OSTMax > 0 Or OthMax > 0 Then
                'KLG is not valid
                ValidationResultInt = 0
                ValidationResultStr = ValidationResultStr & VisitName & " TF KLG 0 may be invalid. "
              Else
                ValidationResultInt = 1
                ValidationResultStr = ValidationResultStr & ""
              End If

            Case "1"

              If JSNMax > 1 Or OSTMax > 1 Then
                'KLG is not valid
                ValidationResultInt = 0
                ValidationResultStr = ValidationResultStr & VisitName & " TF KLG 1 may be invalid. "
              Else
                If JSNMax < 1 And OSTMax < 1 Then
                  'KLG is not valid
                  ValidationResultInt = 0
                  ValidationResultStr = ValidationResultStr & VisitName & " TF KLG 1 may be invalid. "
                Else
                  ValidationResultInt = 1
                  ValidationResultStr = ValidationResultStr & ""
                End If
              End If

            Case "1.9"

              If JSNMax > 0 Or OstMin < 1 Then
                'KLG is not valid
                ValidationResultInt = 0
                ValidationResultStr = ValidationResultStr & VisitName & " TF KLG 2N may be invalid. "
              Else
                ValidationResultInt = 1
                ValidationResultStr = ValidationResultStr & ""
              End If

            Case "2"

              If JSNMax > 1 Or OstMax > 3 Or OstMin < 1 Then
                'KLG is not valid
                ValidationResultInt = 0
                ValidationResultStr = ValidationResultStr & VisitName & " TF KLG 2 may be invalid. "
              Else
                ValidationResultInt = 1
                ValidationResultStr = ValidationResultStr & ""
              End If

            Case "3"

              If JSNMax > 2 Or OstMax > 3 Then
                'KLG is not valid
                ValidationResultInt = 0
                ValidationResultStr = ValidationResultStr & VisitName & " TF KLG 3 may be invalid. "
              Else
                ValidationResultInt = 1
                ValidationResultStr = ValidationResultStr & ""
              End If

            Case "4"

              If JSNMax > 3 Or JSNMin < 2 Or OstMax > 3 Then
                'KLG is not valid
                ValidationResultInt = 0
                ValidationResultStr = ValidationResultStr & VisitName & " TF KLG 4 may be invalid. "
              Else
                ValidationResultInt = 1
                ValidationResultStr = ValidationResultStr & ""
              End If

          End Select

        Else

          ValidationResultInt = 2
          ValidationResultStr = ValidationResultStr & ""

        End If 'JSN/OST not null

      Else 'Output indicator for special missing value

        ValidationResultInt = 2
        ValidationResultStr = ValidationResultStr & ""

      End If 'missing value

    Else 'Output indicator for empty value

      ValidationResultInt = 2
      ValidationResultStr = ValidationResultStr & ""

    End If 'empty

    'Return checks as item and keys in referenced collection object
    ValidationResult.Add item := ValidationResultInt, key := ValidationKeyInt
    ValidationResult.Add item := ValidationResultStr, key := ValidationKeyStr

  End If 'combo locked

End Function

Public Function MOST_Validate_PA_2N_Invalid(ReadingIDIn As String, VisitStrIn As String, SideView As String, ByRef ValidationResult As Collection)

  Dim TableName As String
  Dim KLGVarName As String
  Dim FilterName1 As String
  Dim FilterName2 As String
  Dim KLGValue As String

  Dim TableName2 As String
  Dim VisitVarName As String
  Dim VisitName As String

  Dim FormName As String
  Dim SubFormControlName As String
  Dim KLGCombo As String
  Dim ComboVisible As Boolean
  Dim ComboUnlocked As Boolean

  Dim ValidationResultInt As Integer
  Dim ValidationResultStr As String
  Dim ValidationKeyInt As String
  Dim ValidationKeyStr As String

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

  TableName2 = "tblReadings"
  VisitVarName = "RV" & VisitStrIn & "TP"
  VisitName = MyLookup(TableName2, VisitVarName, FilterName1, ReadingIDIn)

  FormName = "Form_MOST_144_168"
  SubFormControlName = "Subform_PA"

  ValidationKeyInt = "RV" & VisitStrIn & SideView & "Int"
  ValidationKeyStr = "RV" & VisitStrIn & SideView & "Str"

  'Initialize validation result variables
  ValidationResultInt = 0
  ValidationResultStr = ""

  'Check if combobox is visible & unlocked
  KLGCombo = "Combo_" & "RV" & VisitStrIn & KLGVarName

  ComboVisible = Forms(FormName).Controls(SubFormControlName).Form.Controls(KLGCombo).Visible
  ComboUnlocked = Not Forms(FormName).Controls(SubFormControlName).Form.Controls(KLGCombo).Locked

  'Continue if combo box is unlocked and visible
  If ComboVisible And ComboUnlocked Then

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

        JSNMax = Nz(JSNMax,"")
        JSNMin = Nz(JSNMin,"")
        OSTMax = Nz(OSTMax,"")
        OSTMin = Nz(OSTMin,"")
        OthMax = Nz(OthMax,0)
        OthMin = Nz(OthMin,0)

        If JSNMax <> "" And JSNMin <> "" And OSTMax <> "" And OSTMin <> "" Then

          JSNMax = CInt(JSNMax)
          JSNMin = CInt(JSNMin)
          OSTMax = CInt(OSTMax)
          OSTMin = CInt(OSTMin)
          OthMax = CInt(OthMax)
          OthMin = CInt(OthMin)

          'Continue to triage based on KLG value if not special missing value
          Select Case KLGValue

            Case "0"

              If JSNMax > 0 Or OSTMax > 0 Or OthMax > 0 Then
                'KLG is not valid
                ValidationResultInt = 0
                ValidationResultStr = ValidationResultStr & VisitName & " TF KLG 0 may be invalid. "
              Else
                ValidationResultInt = 1
                ValidationResultStr = ValidationResultStr & ""
              End If

            Case "1"

              If JSNMax > 1 Or OSTMax > 1 Then
                'KLG is not valid
                ValidationResultInt = 0
                ValidationResultStr = ValidationResultStr & VisitName & " TF KLG 1 may be invalid. "
              Else
                If JSNMax < 1 And OSTMax < 1 Then
                  'KLG is not valid
                  ValidationResultInt = 0
                  ValidationResultStr = ValidationResultStr & VisitName & " TF KLG 1 may be invalid. "
                Else
                  ValidationResultInt = 1
                  ValidationResultStr = ValidationResultStr & ""
                End If
              End If

            Case "1.9"

                ValidationResultInt = 0
                ValidationResultStr = ValidationResultStr & VisitName & " TF KLG 2N is invalid for this visit. "

            Case "2"

              If JSNMax > 1 Or OstMax > 3 Or OstMin < 1 Then
                'KLG is not valid
                ValidationResultInt = 0
                ValidationResultStr = ValidationResultStr & VisitName & " TF KLG 2 may be invalid. "
              Else
                ValidationResultInt = 1
                ValidationResultStr = ValidationResultStr & ""
              End If

            Case "3"

              If JSNMax > 2 Or OstMax > 3 Then
                'KLG is not valid
                ValidationResultInt = 0
                ValidationResultStr = ValidationResultStr & VisitName & " TF KLG 3 may be invalid. "
              Else
                ValidationResultInt = 1
                ValidationResultStr = ValidationResultStr & ""
              End If

            Case "4"

              If JSNMax > 3 Or JSNMin < 2 Or OstMax > 3 Then
                'KLG is not valid
                ValidationResultInt = 0
                ValidationResultStr = ValidationResultStr & VisitName & " TF KLG 4 may be invalid. "
              Else
                ValidationResultInt = 1
                ValidationResultStr = ValidationResultStr & ""
              End If

          End Select

        Else

          ValidationResultInt = 2
          ValidationResultStr = ValidationResultStr & ""

        End If 'JSN/OST not null

      Else 'Output indicator for special missing value

        ValidationResultInt = 2
        ValidationResultStr = ValidationResultStr & ""

      End If 'missing value

    Else 'Output indicator for empty value

      ValidationResultInt = 2
      ValidationResultStr = ValidationResultStr & ""

    End If 'empty

    'Return checks as item and keys in referenced collection object
    ValidationResult.Add item := ValidationResultInt, key := ValidationKeyInt
    ValidationResult.Add item := ValidationResultStr, key := ValidationKeyStr

  End If 'combo locked

End Function


Public Function MOST_Validate_LAT_Standard(ReadingIDIn As String, VisitStrIn As String, SideView As String, ByRef ValidationResult As Collection)

  Dim TableName As String
  Dim KLGVarName As String
  Dim FilterName1 As String
  Dim FilterName2 As String
  Dim KLGValue As String

  Dim TableName2 As String
  Dim VisitVarName As String
  Dim VisitName As String

  Dim FormName As String
  Dim SubFormControlName As String
  Dim KLGCombo As String
  Dim ComboVisible As Boolean
  Dim ComboUnlocked As Boolean

  Dim ValidationResultInt As Integer
  Dim ValidationResultStr As String
  Dim ValidationKeyInt As String
  Dim ValidationKeyStr As String

  Dim JSNMax As Variant
  Dim JSNMin As Variant
  Dim OstMax As Variant
  Dim OstMin As Variant
  Dim OthMax As Variant
  Dim OthMin As Variant

  'Preset variables
  TableName = "tblScores"
  KLGVarName = SideView & "PFKLG"
  FilterName1 = "READINGID"
  FilterName2 = "RVNUM"

  TableName2 = "tblReadings"
  VisitVarName = "RV" & VisitStrIn & "TP"
  VisitName = MyLookup(TableName2, VisitVarName, FilterName1, ReadingIDIn)

  FormName = "Form_MOST_144_168"

  If SideView = "LXL" Then
    SubFormControlName = "Subform_LLAT"
  Elseif SideView = "LXR" Then
    SubFormControlName = "Subform_RLAT"
  End If

  ValidationKeyInt = "RV" & VisitStrIn & SideView & "Int"
  ValidationKeyStr = "RV" & VisitStrIn & SideView & "Str"

  'Initialize validation result variables
  ValidationResultInt = 0
  ValidationResultStr = ""

  'Check if combobox is visible & unlocked
  KLGCombo = "Combo_" & "RV" & VisitStrIn & KLGVarName

  ComboVisible = Forms(FormName).Controls(SubFormControlName).Form.Controls(KLGCombo).Visible
  ComboUnlocked = Not Forms(FormName).Controls(SubFormControlName).Form.Controls(KLGCombo).Locked

  'Continue if combo box is unlocked and visible
  If ComboVisible And ComboUnlocked Then

    'Get KLG value
    KLGValue = MyLookup2(TableName, KLGVarName, FilterName1, ReadingIDIn, FilterName2, VisitStrIn)

    'Check if empty
    If Nz(KLGValue,"") <> "" Then

      'Check if special missing value
      If KLGValue <> "-6" And KLGValue <> "-7" And KLGValue <> "-8" And KLGValue <> "-9" Then

        'Calculate feature mins and maxes
        JSNMax = MOST_Validate_Features_Max(ReadingIDIn, VisitStrIn, SideView, MOST_Validation_LAT_JSN_Array)
        JSNMin = MOST_Validate_Features_Min(ReadingIDIn, VisitStrIn, SideView, MOST_Validation_LAT_JSN_Array)
        OSTMax = MOST_Validate_Features_Max(ReadingIDIn, VisitStrIn, SideView, MOST_Validation_LAT_OST_Array)
        OSTMin = MOST_Validate_Features_Min(ReadingIDIn, VisitStrIn, SideView, MOST_Validation_LAT_OST_Array)
        OthMax = MOST_Validate_Features_Max(ReadingIDIn, VisitStrIn, SideView, MOST_Validation_LAT_Other_Array)
        OthMin = MOST_Validate_Features_Min(ReadingIDIn, VisitStrIn, SideView, MOST_Validation_LAT_Other_Array)

        JSNMax = Nz(JSNMax,"")
        JSNMin = Nz(JSNMin,"")
        OSTMax = Nz(OSTMax,"")
        OSTMin = Nz(OSTMin,"")
        OthMax = Nz(OthMax,0)
        OthMin = Nz(OthMin,0)

        If JSNMax <> "" And JSNMin <> "" And OSTMax <> "" And OSTMin <> "" Then

          JSNMax = CInt(JSNMax)
          JSNMin = CInt(JSNMin)
          OSTMax = CInt(OSTMax)
          OSTMin = CInt(OSTMin)
          OthMax = CInt(OthMax)
          OthMin = CInt(OthMin)

          'Continue to triage based on KLG value if not special missing value
          Select Case KLGValue

            Case "0"

              If JSNMax > 0 Or OSTMax > 0 Or OthMax > 0 Then
                'KLG is not valid
                ValidationResultInt = 0
                ValidationResultStr = ValidationResultStr & VisitName & " PF KLG 0 may be invalid. "
              Else
                ValidationResultInt = 1
                ValidationResultStr = ValidationResultStr & ""
              End If

            Case "1"

              If JSNMax > 1 Or OSTMax > 1 Then
                'KLG is not valid
                ValidationResultInt = 0
                ValidationResultStr = ValidationResultStr & VisitName & " PF KLG 1 may be invalid. "
              Else
                If JSNMax < 1 And OSTMax < 1 Then
                  'KLG is not valid
                  ValidationResultInt = 0
                  ValidationResultStr = ValidationResultStr & VisitName & " PF KLG 1 may be invalid. "
                Else
                  ValidationResultInt = 1
                  ValidationResultStr = ValidationResultStr & ""
                End If
              End If

            Case "1.9"

              If JSNMax > 0 Or OstMin < 1 Then
                'KLG is not valid
                ValidationResultInt = 0
                ValidationResultStr = ValidationResultStr & VisitName & " PF KLG 2N may be invalid. "
              Else
                ValidationResultInt = 1
                ValidationResultStr = ValidationResultStr & ""
              End If

            Case "2"

              If JSNMax > 1 Or OstMax > 3 Or OstMin < 1 Then
                'KLG is not valid
                ValidationResultInt = 0
                ValidationResultStr = ValidationResultStr & VisitName & " PF KLG 2 may be invalid. "
              Else
                ValidationResultInt = 1
                ValidationResultStr = ValidationResultStr & ""
              End If

            Case "3"

              If JSNMax > 2 Or OstMax > 3 Then
                'KLG is not valid
                ValidationResultInt = 0
                ValidationResultStr = ValidationResultStr & VisitName & " PF KLG 3 may be invalid. "
              Else
                ValidationResultInt = 1
                ValidationResultStr = ValidationResultStr & ""
              End If

            Case "4"

              If JSNMax > 3 Or JSNMin < 2 Or OstMax > 3 Then
                'KLG is not valid
                ValidationResultInt = 0
                ValidationResultStr = ValidationResultStr & VisitName & " PF KLG 4 may be invalid. "
              Else
                ValidationResultInt = 1
                ValidationResultStr = ValidationResultStr & ""
              End If

          End Select

        Else

          ValidationResultInt = 2
          ValidationResultStr = ValidationResultStr & ""

        End If 'JSN/OST not null

      Else 'Output indicator for special missing value

        ValidationResultInt = 2
        ValidationResultStr = ValidationResultStr & ""

      End If 'missing value

    Else 'Output indicator for empty value

      ValidationResultInt = 2
      ValidationResultStr = ValidationResultStr & ""

    End If 'empty

    'Return checks as item and keys in referenced collection object
    ValidationResult.Add item := ValidationResultInt, key := ValidationKeyInt
    ValidationResult.Add item := ValidationResultStr, key := ValidationKeyStr

  End If 'combo locked

End Function

Public Function MOST_Validate_LAT_2N_Invalid(ReadingIDIn As String, VisitStrIn As String, SideView As String, ByRef ValidationResult As Collection)

  Dim TableName As String
  Dim KLGVarName As String
  Dim FilterName1 As String
  Dim FilterName2 As String
  Dim KLGValue As String

  Dim TableName2 As String
  Dim VisitVarName As String
  Dim VisitName As String

  Dim FormName As String
  Dim SubFormControlName As String
  Dim KLGCombo As String
  Dim ComboVisible As Boolean
  Dim ComboUnlocked As Boolean

  Dim ValidationResultInt As Integer
  Dim ValidationResultStr As String
  Dim ValidationKeyInt As String
  Dim ValidationKeyStr As String

  Dim JSNMax As Variant
  Dim JSNMin As Variant
  Dim OstMax As Variant
  Dim OstMin As Variant
  Dim OthMax As Variant
  Dim OthMin As Variant

  'Preset variables
  TableName = "tblScores"
  KLGVarName = SideView & "PFKLG"
  FilterName1 = "READINGID"
  FilterName2 = "RVNUM"

  TableName2 = "tblReadings"
  VisitVarName = "RV" & VisitStrIn & "TP"
  VisitName = MyLookup(TableName2, VisitVarName, FilterName1, ReadingIDIn)

  FormName = "Form_MOST_144_168"

  If SideView = "LXL" Then
    SubFormControlName = "Subform_LLAT"
  Elseif SideView = "LXR" Then
    SubFormControlName = "Subform_RLAT"
  End If

  ValidationKeyInt = "RV" & VisitStrIn & SideView & "Int"
  ValidationKeyStr = "RV" & VisitStrIn & SideView & "Str"

  'Initialize validation result variables
  ValidationResultInt = 0
  ValidationResultStr = ""

  'Check if combobox is visible & unlocked
  KLGCombo = "Combo_" & "RV" & VisitStrIn & KLGVarName

  ComboVisible = Forms(FormName).Controls(SubFormControlName).Form.Controls(KLGCombo).Visible
  ComboUnlocked = Not Forms(FormName).Controls(SubFormControlName).Form.Controls(KLGCombo).Locked

  'Continue if combo box is unlocked and visible
  If ComboVisible And ComboUnlocked Then

    'Get KLG value
    KLGValue = MyLookup2(TableName, KLGVarName, FilterName1, ReadingIDIn, FilterName2, VisitStrIn)

    'Check if empty
    If Nz(KLGValue,"") <> "" Then

      'Check if special missing value
      If KLGValue <> "-6" And KLGValue <> "-7" And KLGValue <> "-8" And KLGValue <> "-9" Then

        'Calculate feature mins and maxes
        JSNMax = MOST_Validate_Features_Max(ReadingIDIn, VisitStrIn, SideView, MOST_Validation_LAT_JSN_Array)
        JSNMin = MOST_Validate_Features_Min(ReadingIDIn, VisitStrIn, SideView, MOST_Validation_LAT_JSN_Array)
        OSTMax = MOST_Validate_Features_Max(ReadingIDIn, VisitStrIn, SideView, MOST_Validation_LAT_OST_Array)
        OSTMin = MOST_Validate_Features_Min(ReadingIDIn, VisitStrIn, SideView, MOST_Validation_LAT_OST_Array)
        OthMax = MOST_Validate_Features_Max(ReadingIDIn, VisitStrIn, SideView, MOST_Validation_LAT_Other_Array)
        OthMin = MOST_Validate_Features_Min(ReadingIDIn, VisitStrIn, SideView, MOST_Validation_LAT_Other_Array)

        JSNMax = Nz(JSNMax,"")
        JSNMin = Nz(JSNMin,"")
        OSTMax = Nz(OSTMax,"")
        OSTMin = Nz(OSTMin,"")
        OthMax = Nz(OthMax,0)
        OthMin = Nz(OthMin,0)

        If JSNMax <> "" And JSNMin <> "" And OSTMax <> "" And OSTMin <> "" Then

          JSNMax = CInt(JSNMax)
          JSNMin = CInt(JSNMin)
          OSTMax = CInt(OSTMax)
          OSTMin = CInt(OSTMin)
          OthMax = CInt(OthMax)
          OthMin = CInt(OthMin)

          'Continue to triage based on KLG value if not special missing value
          Select Case KLGValue

            Case "0"

              If JSNMax > 0 Or OSTMax > 0 Or OthMax > 0 Then
                'KLG is not valid
                ValidationResultInt = 0
                ValidationResultStr = ValidationResultStr & VisitName & " PF KLG 0 may be invalid. "
              Else
                ValidationResultInt = 1
                ValidationResultStr = ValidationResultStr & ""
              End If

            Case "1"

              If JSNMax > 1 Or OSTMax > 1 Then
                'KLG is not valid
                ValidationResultInt = 0
                ValidationResultStr = ValidationResultStr & VisitName & " PF KLG 1 may be invalid. "
              Else
                If JSNMax < 1 And OSTMax < 1 Then
                  'KLG is not valid
                  ValidationResultInt = 0
                  ValidationResultStr = ValidationResultStr & VisitName & " PF KLG 1 may be invalid. "
                Else
                  ValidationResultInt = 1
                  ValidationResultStr = ValidationResultStr & ""
                End If
              End If

            Case "1.9"

                'KLG is not valid
                ValidationResultInt = 0
                ValidationResultStr = ValidationResultStr & VisitName & " TF KLG 2N is invalid for this visit. "

            Case "2"

              If JSNMax > 1 Or OstMax > 3 Or OstMin < 1 Then
                'KLG is not valid
                ValidationResultInt = 0
                ValidationResultStr = ValidationResultStr & VisitName & " PF KLG 2 may be invalid. "
              Else
                ValidationResultInt = 1
                ValidationResultStr = ValidationResultStr & ""
              End If

            Case "3"

              If JSNMax > 2 Or OstMax > 3 Then
                'KLG is not valid
                ValidationResultInt = 0
                ValidationResultStr = ValidationResultStr & VisitName & " PF KLG 3 may be invalid. "
              Else
                ValidationResultInt = 1
                ValidationResultStr = ValidationResultStr & ""
              End If

            Case "4"

              If JSNMax > 3 Or JSNMin < 2 Or OstMax > 3 Then
                'KLG is not valid
                ValidationResultInt = 0
                ValidationResultStr = ValidationResultStr & VisitName & " PF KLG 4 may be invalid. "
              Else
                ValidationResultInt = 1
                ValidationResultStr = ValidationResultStr & ""
              End If

          End Select

        Else

          ValidationResultInt = 2
          ValidationResultStr = ValidationResultStr & ""

        End If 'JSN/OST not null

      Else 'Output indicator for special missing value

        ValidationResultInt = 2
        ValidationResultStr = ValidationResultStr & ""

      End If 'missing value

    Else 'Output indicator for empty value

      ValidationResultInt = 2
      ValidationResultStr = ValidationResultStr & ""

    End If 'empty

    'Return checks as item and keys in referenced collection object
    ValidationResult.Add item := ValidationResultInt, key := ValidationKeyInt
    ValidationResult.Add item := ValidationResultStr, key := ValidationKeyStr

  End If 'combo locked

End Function

Public Function MOST_Validate_MsgBox(ByRef ValidationResult As Collection) As Integer

  Dim ValidationFlag As Boolean
  Dim ValidationMessage As String
  Dim ValidationResponse As Integer
  Dim DummyInt As Integer
  Dim DummyStr As String

  'Preset values
  ValidationFlag = False
  ValidationMessage = ""
  ValidationResponse = vbYes

  'Check each validation result for invalid scores

  DummyInt = ValidationResult.Item("RV3XRInt") 'Right PA'
  If DummyInt = 0 Then
    ValidationFlag = True
    DummyStr = ValidationResult.Item("RV3XRStr")
    ValidationMessage = ValidationMessage & DummyStr
  End If
  DummyInt = ValidationResult.Item("RV4XRInt")
  If DummyInt = 0 Then
    ValidationFlag = True
    DummyStr = ValidationResult.Item("RV4XRStr")
    ValidationMessage = ValidationMessage & DummyStr
  End If

  DummyInt = ValidationResult.Item("RV3XLInt") 'Left PA'
  If DummyInt = 0 Then
    ValidationFlag = True
    DummyStr = ValidationResult.Item("RV3XLStr")
    ValidationMessage = ValidationMessage & DummyStr
  End If
  DummyInt = ValidationResult.Item("RV4XLInt")
  If DummyInt = 0 Then
    ValidationFlag = True
    DummyStr = ValidationResult.Item("RV4XLStr")
    ValidationMessage = ValidationMessage & DummyStr
  End If

  DummyInt = ValidationResult.Item("RV3LXRInt") 'Right Lat'
  If DummyInt = 0 Then
    ValidationFlag = True
    DummyStr = ValidationResult.Item("RV3LXRStr")
    ValidationMessage = ValidationMessage & DummyStr
  End If
  DummyInt = ValidationResult.Item("RV4LXRInt")
  If DummyInt = 0 Then
    ValidationFlag = True
    DummyStr = ValidationResult.Item("RV4LXRStr")
    ValidationMessage = ValidationMessage & DummyStr
  End If

  DummyInt = ValidationResult.Item("RV3LXLInt") 'Left Lat'
  If DummyInt = 0 Then
    ValidationFlag = True
    DummyStr = ValidationResult.Item("RV3LXLStr")
    ValidationMessage = ValidationMessage & DummyStr
  End If
  DummyInt = ValidationResult.Item("RV4LXLInt")
  If DummyInt = 0 Then
    ValidationFlag = True
    DummyStr = ValidationResult.Item("RV4LXLStr")
    ValidationMessage = ValidationMessage & DummyStr
  End If

  'Pop up MsgBox if any invalid scores
  If ValidationFlag Then
    ValidationResponse = MsgBox("Warning: " & vbCrLf & ValidationMessage & "Do you have want to save/sign anyways?", vbYesNo + vbCritical + vbDefaultButton2, "Confirm")
  Else
    ValidationResponse = vbYes
  End If

  MOST_Validate_MsgBox = ValidationResponse

End Function
