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

Public Function MOST_Load_SelectStr()

  SelectStr_TFKLG = "SELECT [DisplayOrder], [ValueStr], [DisplayStr], [ValueDescription] FROM tblValuesTFKLG;"
  SelectStr_PFKLG = "SELECT [DisplayOrder], [ValueStr], [DisplayStr], [ValueDescription] FROM tblValuesPFKLG;"
  SelectStr_JSN = "SELECT [DisplayOrder], [ValueStr], [DisplayStr], [ValueDescription] FROM tblValuesJSN;"
  SelectStr_OS = "SELECT [DisplayOrder], [ValueStr], [DisplayStr], [ValueDescription] FROM tblValuesOS;"
  SelectStr_TFCyst = "SELECT [DisplayOrder], [ValueStr], [DisplayStr], [ValueDescription] FROM tblValuesTFCyst;"
  SelectStr_PFCyst = "SELECT [DisplayOrder], [ValueStr], [DisplayStr], [ValueDescription] FROM tblValuesPFCyst;"
  SelectStr_Sclerosis = "SELECT [DisplayOrder], [ValueStr], [DisplayStr], [ValueDescription] FROM tblValuesSclerosis;"
  SelectStr_Ossification = "SELECT [DisplayOrder], [ValueStr], [DisplayStr], [ValueDescription] FROM tblValuesOssification;"
  SelectStr_MiscYN = "SELECT [DisplayOrder], [ValueStr], [DisplayStr], [ValueDescription] FROM tblValuesMiscYN;"
  
End Function
