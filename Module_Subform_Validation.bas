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

  'Preset variables
  TableName = "tblScores"
  KLGVarName = SideView & "TFKLG"
  FilterName1 = "READINGID"
  FilterName2 = "RVNUM"

  FormName = "Form_MOST_144_168"
  SubFormControlName = "Subform_PA"

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

    'Check if special missing value

    'Continue to triage based on KLG value

  End If

  'Return checks

End Function
