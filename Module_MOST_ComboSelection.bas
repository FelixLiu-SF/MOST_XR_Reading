Attribute VB_Name = "Module_MOST_ComboSelection"
Option Compare Database
Option Explicit

Public SelectStr_TFKLG As String
Public SelectStr_PFKLG As String
Public SelectStr_JSN As String
Public SelectStr_OS As String
Public SelectStr_TFCyst As String
Public SelectStr_PFCyst As String
Public SelectStr_Sclerosis As String
Public SelectStr_Ossification As String
Public SelectStr_MiscYN As String

SelectStr_TFKLG = "SELECT [DisplayOrder], [ValueStr], [DisplayStr], [ValueDescription] FROM tblValuesTFKLG;"
SelectStr_PFKLG = "SELECT [DisplayOrder], [ValueStr], [DisplayStr], [ValueDescription] FROM tblValuesPFKLG;"
SelectStr_JSN = "SELECT [DisplayOrder], [ValueStr], [DisplayStr], [ValueDescription] FROM tblValuesJSN;"
SelectStr_OS = "SELECT [DisplayOrder], [ValueStr], [DisplayStr], [ValueDescription] FROM tblValuesOS;"
SelectStr_TFCyst = "SELECT [DisplayOrder], [ValueStr], [DisplayStr], [ValueDescription] FROM tblValuesTFCyst;"
SelectStr_PFCyst = "SELECT [DisplayOrder], [ValueStr], [DisplayStr], [ValueDescription] FROM tblValuesPFCyst;"
SelectStr_Sclerosis = "SELECT [DisplayOrder], [ValueStr], [DisplayStr], [ValueDescription] FROM tblValuesSclerosis;"
SelectStr_Ossification = "SELECT [DisplayOrder], [ValueStr], [DisplayStr], [ValueDescription] FROM tblValuesOssification;"
SelectStr_MiscYN = "SELECT [DisplayOrder], [ValueStr], [DisplayStr], [ValueDescription] FROM tblValuesMiscYN;"
