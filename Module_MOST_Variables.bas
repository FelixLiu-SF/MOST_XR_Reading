Attribute VB_Name = "Module_MOST_Variables"
Option Compare Database
Option Explicit

Global MOST_Visits_Array(4) As String
Global MOST_RV12_Array(2) As String
Global MOST_RV34_Array(2) As String
Global MOST_PAKnee_Array(2) As String
Global MOST_LATKnee_Array(2) As String
Global MOST_PARoot_Array(19) As String
Global MOST_LATRoot_Array(18) As String

Global MOST_RV1234_XB_Vars() As String
Global MOST_RV12_XB_Vars() As String
Global MOST_RV34_XB_Vars() As String

Global MOST_RV1234_LXB_Vars() As String
Global MOST_RV12_LXB_Vars() As String
Global MOST_RV34_LXB_Vars() As String

Global MOST_RV12_LXR_Vars() As String
Global MOST_RV12_LXL_Vars() As String
Global MOST_RV34_LXR_Vars() As String
Global MOST_RV34_LXL_Vars() As String

Public Function MOST_Load_VariableNameArrays()

  Dim MOST_RLATKnee_Vars() As String
  Dim MOST_LLATKnee_Vars() As String

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

  MOST_RLATKnee_Vars = Concat_Prefix("LXR",MOST_LATRoot_Array)
  MOST_LLATKnee_Vars = Concat_Prefix("LXL",MOST_LATRoot_Array)

  MOST_RV1234_XB_Vars = Concat_VisitVarSide(MOST_Visits_Array, MOST_PAKnee_Array, MOST_PARoot_Array)
  MOST_RV12_XB_Vars = Concat_VisitVarSide(MOST_RV12_Array, MOST_PAKnee_Array, MOST_PARoot_Array)
  MOST_RV34_XB_Vars = Concat_VisitVarSide(MOST_RV34_Array, MOST_PAKnee_Array, MOST_PARoot_Array)

  MOST_RV1234_LXB_Vars = Concat_VisitVarSide(MOST_Visits_Array, MOST_LATKnee_Array, MOST_LATRoot_Array)
  MOST_RV12_LXB_Vars = Concat_VisitVarSide(MOST_RV12_Array, MOST_LATKnee_Array, MOST_LATRoot_Array)
  MOST_RV34_LXB_Vars = Concat_VisitVarSide(MOST_RV34_Array, MOST_LATKnee_Array, MOST_LATRoot_Array)

  MOST_RV12_LXR_Vars = Concat_VisitVar(MOST_RV12_Array, MOST_RLATKnee_Vars)
  MOST_RV12_LXL_Vars = Concat_VisitVar(MOST_RV12_Array, MOST_LLATKnee_Vars)

  MOST_RV34_LXR_Vars = Concat_VisitVar(MOST_RV34_Array, MOST_RLATKnee_Vars)
  MOST_RV34_LXL_Vars = Concat_VisitVar(MOST_RV34_Array, MOST_LLATKnee_Vars)

End Function