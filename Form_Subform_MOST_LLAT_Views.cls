VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Subform_MOST_LLAT_Views"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Current()

    'Load visit text
    
    'Load visit study date
    
    'Load ComboBox Values
    LoadValues_LLat ("LXL")
    
End Sub

'---LOADVALUES_LLAT---'
Private Function LoadValues_LLat(ViewPrefix As String)
'Set table values for ComboBox controls on Left Lateral view subform

    Dim DummyBoolean As Boolean
    Dim TableName As String
    Dim FilterName1 As String
    Dim FilterValue1 As String
    
    'Default filter values
    TableName = "tblScores"
    FilterName1 = "READINGID"
    
    'Get READINGID from current record
    FilterValue1 = Forms("Form_MOST_144_168").Recordset.Fields("READINGID").Value

    'Error catching
    On Error GoTo ErrorHandler_Main1

    'Set menu for PF KLG
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "PFKLG", TableName, FilterName1, FilterValue1)

    'Set menu for PF JSN
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "PFJSN", TableName, FilterName1, FilterValue1)

    'Set menu for "FT" JSN from lateral view
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "FTJSM", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "FTJSL", TableName, FilterName1, FilterValue1)

    'Set menu for Osteophytes
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "OSFA", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "OSFP", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "OSPS", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "OSPI", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "OSTA", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "OSTP", TableName, FilterName1, FilterValue1)

    'Set menu for Sclerosis
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "SCPF", TableName, FilterName1, FilterValue1)

    'Set menu for Cysts
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "CYPF", TableName, FilterName1, FilterValue1)

    'Set menu for Chondrocalcinosis
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "CHON", TableName, FilterName1, FilterValue1)

    'Set menu for Joint Effusion
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "JE", TableName, FilterName1, FilterValue1)

    'Set menu for Ossification
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "OSQI", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "OPTU", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "OPTL", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "OSLB", TableName, FilterName1, FilterValue1)


    'Clear the error catching
    On Error GoTo 0

    Exit Function

ErrorHandler_Main1:
    Resume Next

End Function
