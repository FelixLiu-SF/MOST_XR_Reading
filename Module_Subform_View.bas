Attribute VB_Name = "Module_Subform_View"
Option Compare Database
Option Explicit


'---LOADVALUES_PA---'
Public Function LoadValues_PA(ViewPrefix As String)
'Set table values for ComboBox controls on PA view subform

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

    'Set menu for TF KLG
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "TFKLG", TableName, FilterName1, FilterValue1)

    'Set menu for TF JSN
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "TFJSM", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "TFJSL", TableName, FilterName1, FilterValue1)

    'Set menu for Osteophytes
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "OSFM", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "OSFL", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "OSTM", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "OSTL", TableName, FilterName1, FilterValue1)

    'Set menu for Sclerosis
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "SCFM", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "SCFL", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "SCTM", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "SCTL", TableName, FilterName1, FilterValue1)

    'Set menu for Cysts
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "CYFM", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "CYFL", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "CYTM", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "CYTL", TableName, FilterName1, FilterValue1)

    'Set menu for Attrition
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "ATTM", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "ATTL", TableName, FilterName1, FilterValue1)


    'Set menu for Chondrocalcinosis
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "CHOM", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "CHOL", TableName, FilterName1, FilterValue1)

    'Clear the error catching
    On Error GoTo 0

    Exit Function

ErrorHandler_Main1:
    Resume Next

End Function

'---LOADVALUES_LAT---'
Public Function LoadValues_Lat(Subform_Lat_Name As String, ViewPrefix As String)
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

    'Set value for PF KLG
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", Subform_Lat_Name, ViewPrefix, "PFKLG", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", Subform_Lat_Name, ViewPrefix, "PFOA", TableName, FilterName1, FilterValue1)

    'Set value for PF JSN
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", Subform_Lat_Name, ViewPrefix, "PFJSN", TableName, FilterName1, FilterValue1)

    'Set value for "FT" JSN from lateral view
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", Subform_Lat_Name, ViewPrefix, "FTJSM", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", Subform_Lat_Name, ViewPrefix, "FTJSL", TableName, FilterName1, FilterValue1)

    'Set value for Osteophytes
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", Subform_Lat_Name, ViewPrefix, "OSFA", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", Subform_Lat_Name, ViewPrefix, "OSFP", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", Subform_Lat_Name, ViewPrefix, "OSPS", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", Subform_Lat_Name, ViewPrefix, "OSPI", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", Subform_Lat_Name, ViewPrefix, "OSTA", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", Subform_Lat_Name, ViewPrefix, "OSTP", TableName, FilterName1, FilterValue1)

    'Set value for Sclerosis
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", Subform_Lat_Name, ViewPrefix, "SCPF", TableName, FilterName1, FilterValue1)

    'Set value for Cysts
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", Subform_Lat_Name, ViewPrefix, "CYPF", TableName, FilterName1, FilterValue1)

    'Set value for Chondrocalcinosis
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", Subform_Lat_Name, ViewPrefix, "CHON", TableName, FilterName1, FilterValue1)

    'Set value for Joint Effusion
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", Subform_Lat_Name, ViewPrefix, "JE", TableName, FilterName1, FilterValue1)

    'Set value for Ossification
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", Subform_Lat_Name, ViewPrefix, "OSQI", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", Subform_Lat_Name, ViewPrefix, "OPTU", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", Subform_Lat_Name, ViewPrefix, "OPTL", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", Subform_Lat_Name, ViewPrefix, "OSLB", TableName, FilterName1, FilterValue1)

    'Set value for PF OA
    DummyBoolean = SetComboValue_RV1234("Form_MOST_144_168", Subform_Lat_Name, ViewPrefix, "PFOA", TableName, FilterName1, FilterValue1)

    'Clear the error catching
    On Error GoTo 0

    Exit Function

ErrorHandler_Main1:
    Resume Next

End Function

'---LOADVISITDATES'
Public Function LoadVisitDates(Subform_Name As String)

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
    Forms("Form_MOST_144_168").Controls(Subform_Name).Form.Controls("Label_RV1TP").Caption = VisitStrs(0)
    Forms("Form_MOST_144_168").Controls(Subform_Name).Form.Controls("Label_RV2TP").Caption = VisitStrs(1)
    Forms("Form_MOST_144_168").Controls(Subform_Name).Form.Controls("Label_RV3TP").Caption = VisitStrs(2)
    Forms("Form_MOST_144_168").Controls(Subform_Name).Form.Controls("Label_RV4TP").Caption = VisitStrs(3)

    'Set exam date strings
    Forms("Form_MOST_144_168").Controls(Subform_Name).Form.Controls("Text_RV1DATE").Value = DateStrs(0)
    Forms("Form_MOST_144_168").Controls(Subform_Name).Form.Controls("Text_RV2DATE").Value = DateStrs(1)
    Forms("Form_MOST_144_168").Controls(Subform_Name).Form.Controls("Text_RV3DATE").Value = DateStrs(2)
    Forms("Form_MOST_144_168").Controls(Subform_Name).Form.Controls("Text_RV4DATE").Value = DateStrs(3)

End Function
