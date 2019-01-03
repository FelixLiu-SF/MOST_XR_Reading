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


'---FORM_LOAD---'
Private Sub Form_Load()
' Load various modules when this scoresheet is opened

    Dim DummyBoolean As Boolean

    'Error catching
    On Error GoTo ErrorHandler_Main1

    'Load the form view and objects
    LoadForm ("Form_MOST_144_168")

    'Load the EFilm Automation object to link with EFilm Workstation
    LoadEFilmAuto

    'Load MOST global variable names and SQL selection strings
    DummyBoolean = MOST_Load_SelectStr
    DummyBoolean = MOST_Load_VariableNameArrays

    'Load default view for ComboBox Controls
    LoadView

    'Set all ComboBox OnLostFocus properties to color coding
    SetComboLostFocus

    'Set all Combobox OnFocus properties to dropdown menus
    LoadDropDown

    'Load initial background color scheme for ComboBox Controls
    LoadColor

    'Load the record numbers
    DummyBoolean = LoadRecNav("Form_MOST_144_168", "Text_CurrentRecord")
    DummyBoolean = LoadMaxRec("Form_MOST_144_168", "Text_MaxRecord")

    'Clear the error catching
    On Error GoTo 0

    Exit Sub

ErrorHandler_Main1:
    Resume Next

End Sub

'---FORM_CURRENT---'
Private Sub Form_Current()
' Reload view for current record/participant ID

    Dim DummyBoolean As Boolean

    'Error catching
    On Error GoTo ErrorHandler_Main1

    'Set the ComboBox After Updates
    LoadAfterUpdate

    'Reload default view for ComboBox Controls
    LoadView

    'Load subform views
    DummyBoolean = LoadVisitDates("Subform_PA")
    DummyBoolean = LoadValues_PA("XR")
    DummyBoolean = LoadValues_PA("XL")
    DummyBoolean = LoadVisitDates("Subform_RLAT")
    DummyBoolean = LoadValues_Lat("Subform_RLAT", "LXR")
    DummyBoolean = LoadVisitDates("Subform_LLAT")
    DummyBoolean = LoadValues_Lat("Subform_LLAT", "LXL")

    'Insert assigned reader into reader ComboBox
    Insert_Reader

    'Reload initial background color scheme for ComboBox Controls
    LoadColor

    'Reload the current record number
    DummyBoolean = LoadRecNav("Form_MOST_144_168", "Text_CurrentRecord")

    'Load the PA films and set subform focus automatically
    Command_LOAD_PA_Click

'Clear the error catching
    On Error GoTo 0

    Exit Sub

ErrorHandler_Main1:
    Resume Next

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

'---LOADCOLOR---'
Private Sub LoadColor()
' Reload all of the combo box background colors

    Dim Index As Integer
    Dim Upper As Integer
    Dim ControlName As String
    Dim DummyBoolean As Boolean

    'Error catching
    On Error GoTo ErrorHandler_Main1

    'Loop through preloaded PA view variable name array on PA subform
    Index = 0
    Upper = UBound(MOST_RV1234_XB_Vars, 1) - 1

    For Index = 0 To Upper
        ControlName = "Combo_" & MOST_RV1234_XB_Vars(Index)
        DummyBoolean = BackcolorCode("Form_MOST_144_168", "Subform_PA", ControlName)
    Next

    'Loop through preloaded Lateral View variable name array on Right Lat subform
    Index = 0
    Upper = UBound(MOST_RV1234_LXB_Vars, 1) - 1

    For Index = 0 To Upper
        ControlName = "Combo_" & MOST_RV1234_LXB_Vars(Index)
        DummyBoolean = BackcolorCode("Form_MOST_144_168", "Subform_RLAT", ControlName)
    Next

    'Loop through preloaded Lateral View variable name array on Left Lat subform
    Index = 0
    Upper = UBound(MOST_RV1234_LXB_Vars, 1) - 1

    For Index = 0 To Upper
        ControlName = "Combo_" & MOST_RV1234_LXB_Vars(Index)
        DummyBoolean = BackcolorCode("Form_MOST_144_168", "Subform_LLAT", ControlName)
    Next

    'Clear the error catching
    On Error GoTo 0

    Exit Sub

ErrorHandler_Main1:
    Resume Next

End Sub

'---INSERT_READER---'
Private Sub Insert_Reader()
  'Insert the selected reader into the reader var
  Dim CUR_READER As String
  Dim Reader_Var As String
  Dim Assigned_Reader As String

  On Error GoTo InsertErr
  CUR_READER = Me("Combo_HeaderREADER").Value
  Reader_Var = Me.Recordset.Fields("READER").Value
  Assigned_Reader = DLookup("AssignedReader", "tblProperties", "RecordID = 1")

  'Automatically fill Combo_HeaderREADER if AssignedReader is not null
  If Len(Nz(Assigned_Reader, "")) > 0 And Len(Nz(CUR_READER, "")) < 1 Then
      Me("Combo_HeaderREADER").Value = Assigned_Reader
      CUR_READER = Assigned_Reader
  End If

  'Insert only if selected reader and is not null, and reader var is
  If Len(Nz(CUR_READER, "")) > 0 And Len(Nz(Reader_Var, "")) < 1 Then
      Me.Recordset.Edit
      Me.Recordset.Fields("READER").Value = CUR_READER
      Me.Recordset.Update
  End If


On Error GoTo 0
Exit Sub

InsertErr:
Resume Next

End Sub

'---COMMAND_NEXT_CLICK---'
Private Sub Command_Next_Click()

    Dim DummyBoolean As Boolean

    DummyBoolean = ButtonNext("Form_MOST_144_168", "E_DATE")

End Sub

'---COMMAND_PREV_CLICK---'
Private Sub Command_Prev_Click()

    Dim DummyBoolean As Boolean

    DummyBoolean = ButtonPrev("Form_MOST_144_168", "E_DATE")

End Sub


'---COMMAND_QUIT_CLICK---'
Private Sub Command_QUIT_Click()

  Dim QuitBoolean As Boolean

  On Error GoTo QuitError

  QuitBoolean = QuitRequest("Form_MOST_144_168", "E_DATE")

  If QuitBoolean Then
    DoCmd.Quit
  End If

  On Error GoTo 0
  Exit Sub

QuitError:
  DoCmd.Quit
  Exit Sub

End Sub
