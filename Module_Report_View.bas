Attribute VB_Name = "Module_Report_View"
Option Compare Database
Option Explicit


'---LOADREPORT---'
Public Function LoadReport(ReportName As String)

    Dim Debug_Flag As Integer

    On Error GoTo ErrorHandler1

    Debug_Flag = DLookup("DebugFlag","tblProperties","RecordID = 1")

    If Debug_Flag < 1 Then

        'Set Focus on Report
        Reports(ReportName).SetFocus

        'Maximize the form window
        DoCmd.Maximize

    End If

    'clear error object
    On Error GoTo -1

    On Error Goto ErrorHandler2

    On Error GoTo 0
    Exit Function

    ErrorHandler1:

        'clear error object
        On Error GoTo -1

        On Error Goto ErrorHandler2

        'Maximize the form window
        DoCmd.Maximize

        Resume Next

    ErrorHandler2:

        Exit Function

End Function

'---LOADCURRECREPORT---'
Public Function LoadCurRecReport(FormName As String, ReportName As String, TextBoxName As String)
'Load report text box label to show current record number

    Dim nCurRec As Integer

    On Error GoTo ErrorHandler1

      'Get current record number and update text box
      nCurRec = Forms(FormName).CurrentRecord
      Reports(ReportName).Controls(TextBoxName).Caption = CStr(nCurRec)

      Exit Function

    ErrorHandler1:
    Exit Function

End Function

'---LOADMAXRECREPORT---'
Public Function LoadMaxRecReport(ReportName As String, TextBoxName As String)
'Load report text box label to show max record number

    Dim nMaxRec As Integer

    On Error GoTo ErrorHandler1

      'Get max record number and update text box
      nMaxRec = DLookup("MaxRecord","tblProperties","RecordID = 1")
      Reports(ReportName).Controls(TextBoxName).Caption = CStr(nMaxRec)

      Exit Function

    ErrorHandler1:
    Exit Function

End Function

'---LOADREPORTVALUES_PA---'
Public Function LoadReportValues_PA(Subreport_PA_Name As String, ViewPrefix As String)
'Set table values for ComboBox controls on PA view subform

    Dim DummyBoolean As Boolean
    Dim TableName As String
    Dim FilterName1 As String
    Dim FilterValue1 As String
    Dim ReportFilter As String

    'Default filter values
    TableName = "tblScores"
    FilterName1 = "READINGID"

    'Error catching
    On Error GoTo ErrorHandler_Main1

    'Get passed filter for report
    ReportFilter = Reports("Report_MOST_144_168").Filter

    'Get READINGID from current record
    FilterValue1 = DLookup(FilterName1, "tblReadings", ReportFilter)

    'Set menu for TF KLG
    DummyBoolean = SetReportValue_RV1234("Report_MOST_144_168", Subreport_PA_Name, ViewPrefix, "TFKLG", TableName, FilterName1, FilterValue1)

    'Set menu for TF JSN
    DummyBoolean = SetReportValue_RV1234("Report_MOST_144_168", Subreport_PA_Name, ViewPrefix, "TFJSM", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetReportValue_RV1234("Report_MOST_144_168", Subreport_PA_Name, ViewPrefix, "TFJSL", TableName, FilterName1, FilterValue1)

    'Set menu for Osteophytes
    DummyBoolean = SetReportValue_RV1234("Report_MOST_144_168", Subreport_PA_Name, ViewPrefix, "OSFM", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetReportValue_RV1234("Report_MOST_144_168", Subreport_PA_Name, ViewPrefix, "OSFL", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetReportValue_RV1234("Report_MOST_144_168", Subreport_PA_Name, ViewPrefix, "OSTM", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetReportValue_RV1234("Report_MOST_144_168", Subreport_PA_Name, ViewPrefix, "OSTL", TableName, FilterName1, FilterValue1)

    'Set menu for Sclerosis
    DummyBoolean = SetReportValue_RV1234("Report_MOST_144_168", Subreport_PA_Name, ViewPrefix, "SCFM", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetReportValue_RV1234("Report_MOST_144_168", Subreport_PA_Name, ViewPrefix, "SCFL", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetReportValue_RV1234("Report_MOST_144_168", Subreport_PA_Name, ViewPrefix, "SCTM", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetReportValue_RV1234("Report_MOST_144_168", Subreport_PA_Name, ViewPrefix, "SCTL", TableName, FilterName1, FilterValue1)

    'Set menu for Cysts
    DummyBoolean = SetReportValue_RV1234("Report_MOST_144_168", Subreport_PA_Name, ViewPrefix, "CYFM", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetReportValue_RV1234("Report_MOST_144_168", Subreport_PA_Name, ViewPrefix, "CYFL", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetReportValue_RV1234("Report_MOST_144_168", Subreport_PA_Name, ViewPrefix, "CYTM", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetReportValue_RV1234("Report_MOST_144_168", Subreport_PA_Name, ViewPrefix, "CYTL", TableName, FilterName1, FilterValue1)

    'Set menu for Attrition
    DummyBoolean = SetReportValue_RV1234("Report_MOST_144_168", Subreport_PA_Name, ViewPrefix, "ATTM", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetReportValue_RV1234("Report_MOST_144_168", Subreport_PA_Name, ViewPrefix, "ATTL", TableName, FilterName1, FilterValue1)


    'Set menu for Chondrocalcinosis
    DummyBoolean = SetReportValue_RV1234("Report_MOST_144_168", Subreport_PA_Name, ViewPrefix, "CHOM", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetReportValue_RV1234("Report_MOST_144_168", Subreport_PA_Name, ViewPrefix, "CHOL", TableName, FilterName1, FilterValue1)

    'Clear the error catching
    On Error GoTo 0

    Exit Function

ErrorHandler_Main1:
    Resume Next

End Function

'---LOADREPORTVALUES_LAT---'
Public Function LoadReportValues_Lat(Subreport_Lat_Name As String, ViewPrefix As String)
'Set table values for ComboBox controls on Right Lateral view subform

    Dim DummyBoolean As Boolean
    Dim TableName As String
    Dim FilterName1 As String
    Dim FilterValue1 As String
    Dim ReportFilter As String

    'Default filter values
    TableName = "tblScores"
    FilterName1 = "READINGID"

    'Error catching
    On Error GoTo ErrorHandler_Main1

    'Get passed filter for report
    ReportFilter = Reports("Report_MOST_144_168").Filter

    'Get READINGID from current record
    FilterValue1 = DLookup(FilterName1, "tblReadings", ReportFilter)

    'Set value for PF KLG
    DummyBoolean = SetReportValue_RV1234("Report_MOST_144_168", Subreport_Lat_Name, ViewPrefix, "PFKLG", TableName, FilterName1, FilterValue1)


    'Set value for PF JSN
    DummyBoolean = SetReportValue_RV1234("Report_MOST_144_168", Subreport_Lat_Name, ViewPrefix, "PFJSN", TableName, FilterName1, FilterValue1)

    'Set value for "FT" JSN from lateral view
    DummyBoolean = SetReportValue_RV1234("Report_MOST_144_168", Subreport_Lat_Name, ViewPrefix, "FTJSM", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetReportValue_RV1234("Report_MOST_144_168", Subreport_Lat_Name, ViewPrefix, "FTJSL", TableName, FilterName1, FilterValue1)

    'Set value for Osteophytes
    DummyBoolean = SetReportValue_RV1234("Report_MOST_144_168", Subreport_Lat_Name, ViewPrefix, "OSFA", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetReportValue_RV1234("Report_MOST_144_168", Subreport_Lat_Name, ViewPrefix, "OSFP", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetReportValue_RV1234("Report_MOST_144_168", Subreport_Lat_Name, ViewPrefix, "OSPS", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetReportValue_RV1234("Report_MOST_144_168", Subreport_Lat_Name, ViewPrefix, "OSPI", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetReportValue_RV1234("Report_MOST_144_168", Subreport_Lat_Name, ViewPrefix, "OSTA", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetReportValue_RV1234("Report_MOST_144_168", Subreport_Lat_Name, ViewPrefix, "OSTP", TableName, FilterName1, FilterValue1)

    'Set value for Sclerosis
    DummyBoolean = SetReportValue_RV1234("Report_MOST_144_168", Subreport_Lat_Name, ViewPrefix, "SCPF", TableName, FilterName1, FilterValue1)

    'Set value for Cysts
    DummyBoolean = SetReportValue_RV1234("Report_MOST_144_168", Subreport_Lat_Name, ViewPrefix, "CYPF", TableName, FilterName1, FilterValue1)

    'Set value for Chondrocalcinosis
    DummyBoolean = SetReportValue_RV1234("Report_MOST_144_168", Subreport_Lat_Name, ViewPrefix, "CHON", TableName, FilterName1, FilterValue1)

    'Set value for Joint Effusion
    DummyBoolean = SetReportValue_RV1234("Report_MOST_144_168", Subreport_Lat_Name, ViewPrefix, "JE", TableName, FilterName1, FilterValue1)

    'Set value for Ossification
    DummyBoolean = SetReportValue_RV1234("Report_MOST_144_168", Subreport_Lat_Name, ViewPrefix, "OSQI", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetReportValue_RV1234("Report_MOST_144_168", Subreport_Lat_Name, ViewPrefix, "OPTU", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetReportValue_RV1234("Report_MOST_144_168", Subreport_Lat_Name, ViewPrefix, "OPTL", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetReportValue_RV1234("Report_MOST_144_168", Subreport_Lat_Name, ViewPrefix, "OSLB", TableName, FilterName1, FilterValue1)

    'Clear the error catching
    On Error GoTo 0

    Exit Function

ErrorHandler_Main1:
    Resume Next

End Function

'---LOADREPORTVISITDATES'
Public Function LoadReportVisitDates(Subreport_Name As String)

    Dim VisitStrs(4) As String
    Dim DateStrs(4) As String
    Dim ReportFilter As String

    ReportFilter = Reports("Report_MOST_144_168").Filter

    'Get time points from current record
    VisitStrs(0) = DLookup("RV1TP", "tblReadings", ReportFilter)
    VisitStrs(1) = DLookup("RV2TP", "tblReadings", ReportFilter)
    VisitStrs(2) = DLookup("RV3TP", "tblReadings", ReportFilter)
    VisitStrs(3) = DLookup("RV4TP", "tblReadings", ReportFilter)

    'Get exam dates from current record
    DateStrs(0) = DLookup("RV1DATE", "tblReadings", ReportFilter)
    DateStrs(1) = DLookup("RV2DATE", "tblReadings", ReportFilter)
    DateStrs(2) = DLookup("RV3DATE", "tblReadings", ReportFilter)
    DateStrs(3) = DLookup("RV4DATE", "tblReadings", ReportFilter)

    'Set time point strings
    Reports("Report_MOST_144_168").Controls(Subreport_Name).Report.Controls("Label_RV1TP").Caption = VisitStrs(0)
    Reports("Report_MOST_144_168").Controls(Subreport_Name).Report.Controls("Label_RV2TP").Caption = VisitStrs(1)
    Reports("Report_MOST_144_168").Controls(Subreport_Name).Report.Controls("Label_RV3TP").Caption = VisitStrs(2)
    Reports("Report_MOST_144_168").Controls(Subreport_Name).Report.Controls("Label_RV4TP").Caption = VisitStrs(3)

    'Set exam date strings
    Reports("Report_MOST_144_168").Controls(Subreport_Name).Report.Controls("Text_RV1DATE").Caption = DateStrs(0)
    Reports("Report_MOST_144_168").Controls(Subreport_Name).Report.Controls("Text_RV2DATE").Caption = DateStrs(1)
    Reports("Report_MOST_144_168").Controls(Subreport_Name).Report.Controls("Text_RV3DATE").Caption = DateStrs(2)
    Reports("Report_MOST_144_168").Controls(Subreport_Name).Report.Controls("Text_RV4DATE").Caption = DateStrs(3)

End Function

'---SETREPORTVALUE_RV1234----'
Public Function SetReportValue_RV1234(ReportName As String, SubReportControlName As String, ViewPrefix As String, VarNameRoot As String, TableName As String, FilterName1 As String, FilterValue1 As String)

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
        DummyBoolean = SetReportValue(ReportName, SubReportControlName, ControlName, VariableName, TableName, FilterName1, FilterValue1, FilterName2, FilterValue2)

    Next

End Function

'---SETFORMVALUE---'
Public Function SetReportValue(ReportName As String, SubReportControlName As String, ControlName As String, VariableName As String, TableName As String, FilterName1 As String, FilterValue1 As String, FilterName2 As String, FilterValue2 As String)
'Read value from MS Access table using filters and insert into report combo box

  Dim TableValue As String

  On Error GoTo ComboValueError

  'Get value from table
  TableValue = Nz(MyLookup2(TableName, VariableName, FilterName1, FilterValue1, FilterName2, FilterValue2),"")

  'Update ComboBox value if value is not null
  If Len(TableValue)>0 Then
    Reports(ReportName).Controls(SubReportControlName).Report.Controls(ControlName).Caption = TableValue
  Else
    Reports(ReportName).Controls(SubReportControlName).Report.Controls(ControlName).Caption = ""
  End If

  On Error GoTo 0
  Exit Function

ComboValueError:
  Resume Next

End Function
