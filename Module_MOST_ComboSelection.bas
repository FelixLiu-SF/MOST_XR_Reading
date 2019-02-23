Attribute VB_Name = "Module_MOST_ComboSelection"
Option Compare Database
Option Explicit

Global SelectStr_TFKLG As String
Global SelectStr_PFKLG As String
Global SelectStr_JSN As String
Global SelectStr_OS As String
Global SelectStr_TFCyst As String
Global SelectStr_PFCyst As String
Global SelectStr_Sclerosis As String
Global SelectStr_Ossification As String
Global SelectStr_MiscYN As String
Global SelectStr_Attrition As String
Global SelectStr_Chondro As String
Global SelectStr_JE As String
Global SelectStr_OssLB As String

'---MOST_LOAD_SELECTSTR---'
Public Function MOST_Load_SelectStr()

  SelectStr_TFKLG = "SELECT [ValueStr], [DisplayStr], [ValueDescription] FROM tblValuesTFKLG;"
  SelectStr_PFKLG = "SELECT [ValueStr], [DisplayStr], [ValueDescription] FROM tblValuesPFKLG;"
  SelectStr_JSN = "SELECT [ValueStr], [DisplayStr], [ValueDescription] FROM tblValuesJSN;"
  SelectStr_OS = "SELECT [ValueStr], [DisplayStr], [ValueDescription] FROM tblValuesOS;"
  SelectStr_TFCyst = "SELECT [ValueStr], [DisplayStr], [ValueDescription] FROM tblValuesTFCyst;"
  SelectStr_PFCyst = "SELECT [ValueStr], [DisplayStr], [ValueDescription] FROM tblValuesPFCyst;"
  SelectStr_Sclerosis = "SELECT [ValueStr], [DisplayStr], [ValueDescription] FROM tblValuesSclerosis;"
  SelectStr_Ossification = "SELECT [ValueStr], [DisplayStr], [ValueDescription] FROM tblValuesOssification;"
  SelectStr_MiscYN = "SELECT [ValueStr], [DisplayStr], [ValueDescription] FROM tblValuesMiscYN;"
  SelectStr_Attrition = "SELECT [ValueStr], [DisplayStr], [ValueDescription] FROM tblValuesAttrition;"
  SelectStr_Chondro = "SELECT [ValueStr], [DisplayStr], [ValueDescription] FROM tblValuesChondro;"
  SelectStr_JE = "SELECT [ValueStr], [DisplayStr], [ValueDescription] FROM tblValuesJE;"
  SelectStr_OssLB = "SELECT [ValueStr], [DisplayStr], [ValueDescription] FROM tblValuesOssLB;"

End Function
