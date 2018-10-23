VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Subform_MOST_RLAT_Views"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Current()

    'Load visit & study date text
    LoadVisitDates
    
    'Load ComboBox Values
    LoadValues_RLat ("LXR")

End Sub

'---LOADVALUES_RLAT---'
Private Function LoadValues_RLat(ViewPrefix As String)
'Set table values for ComboBox controls on Right Lateral view subform

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
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "PFKLG", TableName, FilterName1, FilterValue1)

    'Set menu for PF JSN
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "PFJSN", TableName, FilterName1, FilterValue1)

    'Set menu for "FT" JSN from lateral view
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "FTJSM", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "FTJSL", TableName, FilterName1, FilterValue1)

    'Set menu for Osteophytes
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "OSFA", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "OSFP", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "OSPS", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "OSPI", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "OSTA", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "OSTP", TableName, FilterName1, FilterValue1)

    'Set menu for Sclerosis
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "SCPF", TableName, FilterName1, FilterValue1)

    'Set menu for Cysts
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "CYPF", TableName, FilterName1, FilterValue1)

    'Set menu for Chondrocalcinosis
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "CHON", TableName, FilterName1, FilterValue1)

    'Set menu for Joint Effusion
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "JE", TableName, FilterName1, FilterValue1)

    'Set menu for Ossification
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "OSQI", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "OPTU", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "OPTL", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "OSLB", TableName, FilterName1, FilterValue1)


    'Clear the error catching
    On Error GoTo 0

    Exit Function

ErrorHandler_Main1:
    Resume Next

End Function

Private Sub LoadVisitDates()

    Dim VisitStrs(4) As String
    Dim DateStrs(4) As String

    'Get time points from current record
    VisitStrs(0) = Forms("Form_MOST_144_168").Recordset.Fields("RV1TP").Value
    VisitStrs(1) = Forms("Form_MOST_144_168").Recordset.Fields("RV2TP").Value
    VisitStrs(2) = Forms("Form_MOST_144_168").Recordset.Fields("RV3TP").Value
    VisitStrs(3) = Forms("Form_MOST_144_168").Recordset.Fields("RV4TP").Value
    
    'Get exam dates from current record
    DateStrs(0) = Forms("Form_MOST_144_168").Recordset.Fields("RV1DATE").Value
    DateStrs(1) = Forms("Form_MOST_144_168").Recordset.Fields("RV2DATE").Value
    DateStrs(2) = Forms("Form_MOST_144_168").Recordset.Fields("RV3DATE").Value
    DateStrs(3) = Forms("Form_MOST_144_168").Recordset.Fields("RV4DATE").Value
    
    'Set time point strings
    Forms("Form_MOST_144_168").Controls("Subform_RLAT").Form.Controls("Label_RV1TP").Caption = VisitStrs(0)
    Forms("Form_MOST_144_168").Controls("Subform_RLAT").Form.Controls("Label_RV2TP").Caption = VisitStrs(1)
    Forms("Form_MOST_144_168").Controls("Subform_RLAT").Form.Controls("Label_RV3TP").Caption = VisitStrs(2)
    Forms("Form_MOST_144_168").Controls("Subform_RLAT").Form.Controls("Label_RV4TP").Caption = VisitStrs(3)
    
    'Set exam date strings
    Forms("Form_MOST_144_168").Controls("Subform_RLAT").Form.Controls("Text_RV1DATE").Value = DateStrs(0)
    Forms("Form_MOST_144_168").Controls("Subform_RLAT").Form.Controls("Text_RV2DATE").Value = DateStrs(1)
    Forms("Form_MOST_144_168").Controls("Subform_RLAT").Form.Controls("Text_RV3DATE").Value = DateStrs(2)
    Forms("Form_MOST_144_168").Controls("Subform_RLAT").Form.Controls("Text_RV4DATE").Value = DateStrs(3)

End Sub
