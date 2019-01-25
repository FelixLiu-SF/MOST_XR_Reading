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

Public Function MOST_Validate_By_ID(ReadingIDIn As String)

  'Check if existing or new cohort

  'Validate right knee PA

  'Validate left knee PA

  'Validate right knee lateral

  'Validate left knee lateral

  'Return results

End Function

Public Function MOST_Validate_PA_Standard(ReadingIDIn As String, VisitStrIn As String, SideView As String)

  Dim TableName As String
  Dim KLGVarName As String
  Dim FilterName1 As String
  Dim FilterName2 As String
  Dim KLGValue As String
  Dim ValidationResultStr As String

  'Preset variables
  TableName = "tblScores"
  KLGVarName = SideView & "TFKLG"
  FilterName1 = "READINGID"
  FilterName2 = "RVNUM"

  ValidationResultStr = ""

  'Check if combobox is visible & unlocked

  'Get KLG value
  KLGValue = MyLookup2(TableName, KLGVarName, FilterName1, ReadingIDIn, FilterName2, VisitStrIn)

  'Check if empty

  'Check if special missing value

  'Continue to triage based on KLG value

  'Return checks

End Function
