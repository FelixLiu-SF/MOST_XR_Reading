VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Report_MOST_144_168"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


'---REPORT_LOAD---'
Private Sub Report_Load()
' Load various modules when this scoresheet is opened

    Dim DummyBoolean As Boolean

    'Error catching
    On Error GoTo ErrorHandler_Main1

    'Load the report view and objects
    LoadReport ("Report_MOST_144_168")

    'Load default view for ComboBox Controls
    'LoadView

    'Load subform views
    DummyBoolean = LoadReportVisitDates("Subreport_PA")
    DummyBoolean = LoadReportValues_PA("Subreport_PA", "XR")
    DummyBoolean = LoadReportValues_PA("Subreport_PA", "XL")
    DummyBoolean = LoadReportVisitDates("Subreport_RLAT")
    DummyBoolean = LoadReportValues_Lat("Subreport_RLAT", "LXR")
    DummyBoolean = LoadReportVisitDates("Subreport_LLAT")
    DummyBoolean = LoadReportValues_Lat("Subreport_LLAT", "LXL")

    'Load the record numbers
    DummyBoolean = LoadCurRecReport("Form_MOST_144_168", "Report_MOST_144_168","Text_CurrentRecord")
    DummyBoolean = LoadMaxRecReport("Report_MOST_144_168", "Text_MaxRecord")

    'Clear the error catching
    On Error GoTo 0

    Exit Sub

ErrorHandler_Main1:
    Resume Next

End Sub

'---REPORT_CURRENT---'
Private Sub Report_Current()
' Reload view for current record/participant ID

    Dim DummyBoolean As Boolean

    'Error catching
    On Error GoTo ErrorHandler_Main1

    'Set the ComboBox After Updates
    'LoadAfterUpdate

    'Reload default view for ComboBox Controls
    'LoadView

    'Load subform views
    DummyBoolean = LoadReportVisitDates("Subreport_PA")
    DummyBoolean = LoadReportValues_PA("Subreport_PA", "XR")
    DummyBoolean = LoadReportValues_PA("Subreport_PA", "XL")
    DummyBoolean = LoadReportVisitDates("Subreport_RLAT")
    DummyBoolean = LoadReportValues_Lat("Subreport_RLAT", "LXR")
    DummyBoolean = LoadReportVisitDates("Subreport_LLAT")
    DummyBoolean = LoadReportValues_Lat("Subreport_LLAT", "LXL")

    'Insert assigned reader into reader ComboBox
    Insert_Report_Reader

    'Reload initial background color scheme for ComboBox Controls
    'LoadColor

    'Load the record numbers
    DummyBoolean = LoadCurRecReport("Form_MOST_144_168", "Report_MOST_144_168","Text_CurrentRecord")
    DummyBoolean = LoadMaxRecReport("Report_MOST_144_168", "Text_MaxRecord")

    'Load the PA films and set subform focus automatically
    'Command_LOAD_PA_Click

'Clear the error catching
    On Error GoTo 0

    Exit Sub

ErrorHandler_Main1:
    Resume Next

End Sub

'---INSERT_REPORT_READER---'
Private Sub Insert_Report_Reader()
  'Insert the selected reader into the reader var
  Dim Cur_Reader As String
  Dim Reader_Var As String
  Dim Assigned_Reader As String

  On Error GoTo InsertErr
  Cur_Reader = Me("Combo_HeaderREADER").Value
  Reader_Var = Me.Recordset.Fields("READER").Value
  Assigned_Reader = DLookup("AssignedReader","tblProperties","RecordID = 1")

  'Automatically fill Combo_HeaderREADER if AssignedReader is not null
  If Len(Nz(Assigned_Reader, "")) > 0 And Len(Nz(Cur_Reader, "")) < 1 Then
      Me("Combo_HeaderREADER").Value = Assigned_Reader
      Cur_Reader = Assigned_Reader
  End If

  'Insert reader name from table instead if it exists
  If Len(Nz(Cur_Reader, "")) > 0 And Len(Nz(Reader_Var, "")) > 0 Then
      Me("Combo_HeaderREADER").Value = Reader_Var
      Cur_Reader = Reader_Var
  End If


On Error GoTo 0
Exit Sub

InsertErr:
Resume Next

End Sub

'---COMMAND_QUIT_CLICK---'
Private Sub Command_QUIT_Click()
'Close the printing report

  On Error GoTo QuitError

  DoCmd.Close

  On Error GoTo 0
  Exit Sub

QuitError:
  Exit Sub

End Sub

'---LOADVIEW---'
Private Sub LoadView()
' Show/Hide/Lock/Unlock the ComboBoxes for default view on all subforms

    Dim ControlName_RV12_PA() As String
    Dim ControlName_RV12_RLAT() As String
    Dim ControlName_RV12_LLAT() As String
    Dim ControlName_RV34_RLAT() As String
    Dim ControlName_RV34_LLAT() As String
    Dim ControlName_RV12_PFKLG() As String
    Dim DebugFlag As Integer
    Dim DebugBoolean As Boolean
    Dim DummyBoolean As Boolean

    'Error catching
    On Error GoTo ErrorHandler_Main1

    'Get DebugFlag
    DebugFlag = DLookup("DebugFlag", "tblProperties", "RecordID = 1")
    If Len(Nz(DebugFlag, "")) > 0 Then
        If DebugFlag > 0 Then
            DebugBoolean = False
        Else
            DebugBoolean = True
        End If

    Else
        DebugBoolean = True
    End If

    'Construct ComboBox Control Names from preloaded MOST variable root names
    ControlName_RV12_PA = Concat_Prefix("Combo_", MOST_RV12_XB_Vars)
    ControlName_RV12_RLAT = Concat_Prefix("Combo_", MOST_RV12_LXR_Vars)
    ControlName_RV12_LLAT = Concat_Prefix("Combo_", MOST_RV12_LXL_Vars)
    ControlName_RV34_RLAT = Concat_Prefix("Combo_", MOST_RV34_LXR_Vars)
    ControlName_RV34_LLAT = Concat_Prefix("Combo_", MOST_RV34_LXL_Vars)
    ControlName_RV12_PFKLG = Concat_Prefix("Combo_", MOST_RV12_LXB_PFKLG)

    'Lock the RV1 & RV2 PA variables on PA subform
    DummyBoolean = LockVars("Form_MOST_144_168", "Subform_PA", ControlName_RV12_PA, DebugBoolean)

    'Lock the RIGHT knee RV1 & RV2 variables on RIGHT Lateral subform
    DummyBoolean = LockVars("Form_MOST_144_168", "Subform_RLAT", ControlName_RV12_RLAT, DebugBoolean)

    'Lock the LEFT knee RV1 & RV2 variables on LEFT Lateral subform
    DummyBoolean = LockVars("Form_MOST_144_168", "Subform_LLAT", ControlName_RV12_LLAT, DebugBoolean)

    'Lock & Hide all of the LEFT knee variables on RIGHT Lateral subform
    DummyBoolean = LockVars("Form_MOST_144_168", "Subform_RLAT", ControlName_RV12_LLAT, DebugBoolean)
    DummyBoolean = HideVars("Form_MOST_144_168", "Subform_RLAT", ControlName_RV12_LLAT, DebugBoolean)

    DummyBoolean = LockVars("Form_MOST_144_168", "Subform_RLAT", ControlName_RV34_LLAT, DebugBoolean)
    DummyBoolean = HideVars("Form_MOST_144_168", "Subform_RLAT", ControlName_RV34_LLAT, DebugBoolean)

    'Lock & Hide all of the RIGHT knee variables on LEFT Lateral subform
    DummyBoolean = LockVars("Form_MOST_144_168", "Subform_LLAT", ControlName_RV12_RLAT, DebugBoolean)
    DummyBoolean = HideVars("Form_MOST_144_168", "Subform_LLAT", ControlName_RV12_RLAT, DebugBoolean)

    DummyBoolean = LockVars("Form_MOST_144_168", "Subform_LLAT", ControlName_RV34_RLAT, DebugBoolean)
    DummyBoolean = HideVars("Form_MOST_144_168", "Subform_LLAT", ControlName_RV34_RLAT, DebugBoolean)

    'Lock & Hide all the PF KLG variables from Visit 1 & 2
    DummyBoolean = LockVars("Form_MOST_144_168", "Subform_RLAT", ControlName_RV12_PFKLG, DebugBoolean)
    DummyBoolean = HideVars("Form_MOST_144_168", "Subform_RLAT", ControlName_RV12_PFKLG, DebugBoolean)
    DummyBoolean = LockVars("Form_MOST_144_168", "Subform_LLAT", ControlName_RV12_PFKLG, DebugBoolean)
    DummyBoolean = HideVars("Form_MOST_144_168", "Subform_LLAT", ControlName_RV12_PFKLG, DebugBoolean)

    'Clear the error catching
    On Error GoTo 0

    Exit Sub

ErrorHandler_Main1:
    Resume Next

End Sub
