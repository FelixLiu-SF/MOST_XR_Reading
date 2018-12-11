Attribute VB_Name = "Module_MOST_Variables"
Option Compare Database
Option Explicit

Global MOST_Visits_Array(4) As String
Global MOST_RV12_Array(2) As String
Global MOST_RV34_Array(2) As String
Global MOST_PAKnee_Array(2) As String
Global MOST_LATKnee_Array(2) As String
Global MOST_PARoot_Array(19) As String
Global MOST_LATRoot_Array(19) As String

Global MOST_RV1234_XB_Vars() As String
Global MOST_RV12_XB_Vars() As String
Global MOST_RV34_XB_Vars() As String

Global MOST_RV1234_LXB_Vars() As String
Global MOST_RV12_LXB_Vars() As String
Global MOST_RV34_LXB_Vars() As String
Global MOST_RV12_LXB_PFKLG() As String

Global MOST_RV12_LXR_Vars() As String
Global MOST_RV12_LXL_Vars() As String
Global MOST_RV34_LXR_Vars() As String
Global MOST_RV34_LXL_Vars() As String

Global MOST_RPAKnee_Vars() As String
Global MOST_LPAKnee_Vars() As String
Global MOST_RLATKnee_Vars() As String
Global MOST_LLATKnee_Vars() As String

Public Function MOST_Load_VariableNameArrays()

  MOST_Visits_Array(0) = "RV1"
  MOST_Visits_Array(1) = "RV2"
  MOST_Visits_Array(2) = "RV3"
  MOST_Visits_Array(3) = "RV4"

  MOST_RV12_Array(0) = "RV1"
  MOST_RV12_Array(1) = "RV2"

  MOST_RV34_Array(0) = "RV3"
  MOST_RV34_Array(1) = "RV4"

  MOST_PAKnee_Array(0) = "XR"
  MOST_PAKnee_Array(1) = "XL"

  MOST_LATKnee_Array(0) = "LXR"
  MOST_LATKnee_Array(1) = "LXL"

  MOST_PARoot_Array(0) = "TFKLG"
  MOST_PARoot_Array(1) = "TFJSM"
  MOST_PARoot_Array(2) = "TFJSL"
  MOST_PARoot_Array(3) = "OSFM"
  MOST_PARoot_Array(4) = "OSFL"
  MOST_PARoot_Array(5) = "OSTM"
  MOST_PARoot_Array(6) = "OSTL"
  MOST_PARoot_Array(7) = "SCFM"
  MOST_PARoot_Array(8) = "SCFL"
  MOST_PARoot_Array(9) = "SCTM"
  MOST_PARoot_Array(10) = "SCTL"
  MOST_PARoot_Array(11) = "CYFM"
  MOST_PARoot_Array(12) = "CYFL"
  MOST_PARoot_Array(13) = "CYTM"
  MOST_PARoot_Array(14) = "CYTL"
  MOST_PARoot_Array(15) = "ATTM"
  MOST_PARoot_Array(16) = "ATTL"
  MOST_PARoot_Array(17) = "CHOM"
  MOST_PARoot_Array(18) = "CHOL"

  MOST_LATRoot_Array(0) = "PFKLG"
  MOST_LATRoot_Array(1) = "PFJSN"
  MOST_LATRoot_Array(2) = "FTJSM"
  MOST_LATRoot_Array(3) = "FTJSL"
  MOST_LATRoot_Array(4) = "OSFA"
  MOST_LATRoot_Array(5) = "OSFP"
  MOST_LATRoot_Array(6) = "OSPS"
  MOST_LATRoot_Array(7) = "OSPI"
  MOST_LATRoot_Array(8) = "OSTA"
  MOST_LATRoot_Array(9) = "OSTP"
  MOST_LATRoot_Array(10) = "SCPF"
  MOST_LATRoot_Array(11) = "CYPF"
  MOST_LATRoot_Array(12) = "CHON"
  MOST_LATRoot_Array(13) = "JE"
  MOST_LATRoot_Array(14) = "OSQI"
  MOST_LATRoot_Array(15) = "OPTU"
  MOST_LATRoot_Array(16) = "OPTL"
  MOST_LATRoot_Array(17) = "OSLB"


  MOST_RPAKnee_Vars = Concat_Prefix("XR", MOST_PARoot_Array)
  MOST_LPAKnee_Vars = Concat_Prefix("XL", MOST_PARoot_Array)
  MOST_RLATKnee_Vars = Concat_Prefix("LXR", MOST_LATRoot_Array)
  MOST_LLATKnee_Vars = Concat_Prefix("LXL", MOST_LATRoot_Array)

  MOST_RV1234_XB_Vars = Concat_VisitVarSide(MOST_Visits_Array, MOST_PAKnee_Array, MOST_PARoot_Array)
  MOST_RV12_XB_Vars = Concat_VisitVarSide(MOST_RV12_Array, MOST_PAKnee_Array, MOST_PARoot_Array)
  MOST_RV34_XB_Vars = Concat_VisitVarSide(MOST_RV34_Array, MOST_PAKnee_Array, MOST_PARoot_Array)

  MOST_RV1234_LXB_Vars = Concat_VisitVarSide(MOST_Visits_Array, MOST_LATKnee_Array, MOST_LATRoot_Array)
  MOST_RV12_LXB_Vars = Concat_VisitVarSide(MOST_RV12_Array, MOST_LATKnee_Array, MOST_LATRoot_Array)
  MOST_RV34_LXB_Vars = Concat_VisitVarSide(MOST_RV34_Array, MOST_LATKnee_Array, MOST_LATRoot_Array)

  MOST_RV12_LXB_PFKLG = Concat_VisitVarSide(MOST_RV12_Array, MOST_LATKnee_Array, MOST_LATRoot_Array(0))

  MOST_RV12_LXR_Vars = Concat_VisitVar(MOST_RV12_Array, MOST_RLATKnee_Vars)
  MOST_RV12_LXL_Vars = Concat_VisitVar(MOST_RV12_Array, MOST_LLATKnee_Vars)

  MOST_RV34_LXR_Vars = Concat_VisitVar(MOST_RV34_Array, MOST_RLATKnee_Vars)
  MOST_RV34_LXL_Vars = Concat_VisitVar(MOST_RV34_Array, MOST_LLATKnee_Vars)

End Function

'---SETCOMBOSELECTION_RV1234---'
Public Function SetComboSelection_RV1234(FormName As String, SubFormControlName As String, ViewPrefix As String, VarNameRoot As String, SelectionStr As String, MenuLimitBoolean As Boolean, ColHeaderBoolean As Boolean)

    Dim DummyBoolean As Boolean
    Dim VisitArray(4) As String
    Dim ControlName As String
    Dim OnFocusStr As String
    Dim Index As Integer

    'Define default variables
    VisitArray(0) = "RV1"
    VisitArray(1) = "RV2"
    VisitArray(2) = "RV3"
    VisitArray(3) = "RV4"

    'Loop through visits
    Index = 0
    For Index = 0 To 4

        'Construct ComboBox Control name
        ControlName = "Combo_" & VisitArray(Index) & ViewPrefix & VarNameRoot

        'Construct ComboBox selection string
        OnFocusStr = Make_ControlUpdate_Func(FormName, SubFormControlName, ControlName, SelectionStr, 4, 2, "0; 0; 0.5 in; 2 in", 3, MenuLimitBoolean, ColHeaderBoolean)

        'Set the selection string to the OnFocus property of the ComboBox
        DummyBoolean = Control_Edit_OnFocus(FormName, SubFormControlName, ControlName, OnFocusStr)

    Next

End Function

'---SETCOMBOUPDATE_RV1234----'
Public Function SetComboUpdate_RV1234(FormName As String, SubFormControlName As String, ViewPrefix As String, VarNameRoot As String, TableName As String, FilterName1 As String, FilterValue1 As String)

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
        AfterUpdateStr = Make_ControlAfterUpdate_Func(FormName, SubFormControlName, ControlName, VariableName, TableName, FilterName1, FilterValue1, FilterName2, FilterValue2)

        'Set the after update string to the OnFocus property of the ComboBox
        DummyBoolean = Control_Edit_AfterUpdate(FormName, SubFormControlName, ControlName, AfterUpdateStr)

    Next

End Function

'---SETCOMBOVALUE_RV1234----'
Public Function SetComboValue_RV1234(FormName As String, SubFormControlName As String, ViewPrefix As String, VarNameRoot As String, TableName As String, FilterName1 As String, FilterValue1 As String)

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

        'Insert value from table into ComboBox
        DummyBoolean = SetComboValue(FormName, SubFormControlName, ControlName, VariableName, TableName, FilterName1, FilterValue1, FilterName2, FilterValue2)

    Next

End Function
