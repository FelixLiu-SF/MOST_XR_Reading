VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Form_MOST_144_168"
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
    Dim DebugFlag As Integer
    Dim DebugBoolean As Boolean
    Dim DummyBoolean As Boolean

    'Error catching
    On Error GoTo ErrorHandler_Main1

    'Get AdjudicationFlag
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

'---SETCOMBOLOSTFOCUS---'
Private Sub SetComboLostFocus()
' Set all of the combo box background colors to update after focus is lost

    Dim Index As Integer
    Dim Upper As Integer
    Dim ControlName As String
    Dim DummyBoolean As Boolean
    Dim LostFocusStr As String

    'Error catching
    On Error GoTo ErrorHandler_Main1

    'Loop through preloaded PA view variable name array on PA subform
    Index = 0
    Upper = UBound(MOST_RV1234_XB_Vars, 1) - 1

    For Index = 0 To Upper
        ControlName = "Combo_" & MOST_RV1234_XB_Vars(Index)

        LostFocusStr = Make_ControlColor_Func("Form_MOST_144_168", "Subform_PA", ControlName)

        DummyBoolean = Control_Edit_LostFocus("Form_MOST_144_168", "Subform_PA", ControlName, LostFocusStr)
    Next

    'Loop through preloaded Lateral View variable name array on Right Lat subform
    Index = 0
    Upper = UBound(MOST_RV1234_LXB_Vars, 1) - 1

    For Index = 0 To Upper
        ControlName = "Combo_" & MOST_RV1234_LXB_Vars(Index)

        LostFocusStr = Make_ControlColor_Func("Form_MOST_144_168", "Subform_RLAT", ControlName)

        DummyBoolean = Control_Edit_LostFocus("Form_MOST_144_168", "Subform_RLAT", ControlName, LostFocusStr)
    Next

    'Loop through preloaded Lateral View variable name array on Left Lat subform
    Index = 0
    Upper = UBound(MOST_RV1234_LXB_Vars, 1) - 1

    For Index = 0 To Upper
        ControlName = "Combo_" & MOST_RV1234_LXB_Vars(Index)

        LostFocusStr = Make_ControlColor_Func("Form_MOST_144_168", "Subform_LLAT", ControlName)

        DummyBoolean = Control_Edit_LostFocus("Form_MOST_144_168", "Subform_LLAT", ControlName, LostFocusStr)
    Next

    'Clear the error catching
    On Error GoTo 0

    Exit Sub

ErrorHandler_Main1:
    Resume Next

End Sub

'---LOADDROPDOWN---'
Private Function LoadDropDown()
'Set dropdown menus for all ComboBox controls on all subforms

    Dim DummyBoolean As Boolean
    Dim AdjFlag As Integer
    Dim AdjBoolean As Boolean

    'Error catching
    On Error GoTo ErrorHandler_Main1

    'Get AdjudicationFlag
    AdjFlag = DLookup("AdjudicationFlag", "tblProperties", "RecordID = 1")
    If Len(Nz(AdjFlag, "")) > 0 Then
        If AdjFlag > 0 Then
            AdjBoolean = False
        Else
            AdjBoolean = True
        End If

    Else
        AdjBoolean = True
    End If

    'Set menus on PA view subform & limit selections to menu list
    DummyBoolean = LoadDropDown_PA("XR", AdjBoolean)
    DummyBoolean = LoadDropDown_PA("XL", AdjBoolean)

    'Set menus on Right Lateral view subform & limit selections to menu list
    DummyBoolean = LoadDropDown_RLat("LXR", AdjBoolean)

    'Set menus on Left Lateral view subform & limit selections to menu list
    DummyBoolean = LoadDropDown_LLat("LXL", AdjBoolean)

    'Clear the error catching
    On Error GoTo 0

    Exit Function

ErrorHandler_Main1:
    Resume Next

End Function

'---LOADAFTERUPDATE---'
Private Function LoadAfterUpdate()
'Set dropdown menus for all ComboBox controls on all subforms

    Dim DummyBoolean As Boolean

    'Error catching
    On Error GoTo ErrorHandler_Main1

    'Set menus on PA view subform & limit selections to menu list
    DummyBoolean = LoadAfterUpdate_PA("XR")
    DummyBoolean = LoadAfterUpdate_PA("XL")

    'Set menus on Right Lateral view subform & limit selections to menu list
    DummyBoolean = LoadAfterUpdate_RLat("LXR")

    'Set menus on Left Lateral view subform & limit selections to menu list
    DummyBoolean = LoadAfterUpdate_LLat("LXL")

    'Clear the error catching
    On Error GoTo 0

    Exit Function

ErrorHandler_Main1:
    Resume Next

End Function

'---LOADDROPDOWN_PA---'
Private Function LoadDropDown_PA(ViewPrefix As String, MenuLimitBoolean As Boolean)
'Set dropdown menus for ComboBox controls on PA view subform

    Dim DummyBoolean As Boolean

    'Error catching
    On Error GoTo ErrorHandler_Main1

    'Set menu for TF KLG
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "TFKLG", SelectStr_TFKLG, MenuLimitBoolean)

    'Set menu for TF JSN
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "TFJSM", SelectStr_JSN, MenuLimitBoolean)
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "TFJSL", SelectStr_JSN, MenuLimitBoolean)

    'Set menu for Osteophytes
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "OSFM", SelectStr_OS, MenuLimitBoolean)
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "OSFL", SelectStr_OS, MenuLimitBoolean)
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "OSTM", SelectStr_OS, MenuLimitBoolean)
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "OSTL", SelectStr_OS, MenuLimitBoolean)

    'Set menu for Sclerosis
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "SCFM", SelectStr_Sclerosis, MenuLimitBoolean)
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "SCFL", SelectStr_Sclerosis, MenuLimitBoolean)
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "SCTM", SelectStr_Sclerosis, MenuLimitBoolean)
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "SCTL", SelectStr_Sclerosis, MenuLimitBoolean)

    'Set menu for Cysts
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "CYFM", SelectStr_TFCyst, MenuLimitBoolean)
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "CYFL", SelectStr_TFCyst, MenuLimitBoolean)
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "CYTM", SelectStr_TFCyst, MenuLimitBoolean)
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "CYTL", SelectStr_TFCyst, MenuLimitBoolean)

    'Set menu for Attrition
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "ATTM", SelectStr_MiscYN, MenuLimitBoolean)
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "ATTL", SelectStr_MiscYN, MenuLimitBoolean)

    'Set menu for Chondrocalcinosis
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "CHOM", SelectStr_MiscYN, MenuLimitBoolean)
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "CHOL", SelectStr_MiscYN, MenuLimitBoolean)


    'Clear the error catching
    On Error GoTo 0

    Exit Function

ErrorHandler_Main1:
    Resume Next

End Function

'---LOADDROPDOWN_RLAT---'
Private Function LoadDropDown_RLat(ViewPrefix As String, MenuLimitBoolean As Boolean)
'Set dropdown menus for ComboBox controls on Right Lateral view subform

    Dim DummyBoolean As Boolean

    'Error catching
    On Error GoTo ErrorHandler_Main1

    'Set menu for PF KLG
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "PFKLG", SelectStr_PFKLG, MenuLimitBoolean)
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "PFOA", SelectStr_MiscYN, MenuLimitBoolean)

    'Set menu for PF JSN
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "PFJSN", SelectStr_JSN, MenuLimitBoolean)

    'Set menu for "FT" JSN from lateral view
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "FTJSM", SelectStr_JSN, MenuLimitBoolean)
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "FTJSL", SelectStr_JSN, MenuLimitBoolean)

    'Set menu for Osteophytes
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "OSFA", SelectStr_OS, MenuLimitBoolean)
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "OSFP", SelectStr_OS, MenuLimitBoolean)
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "OSPS", SelectStr_OS, MenuLimitBoolean)
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "OSPI", SelectStr_OS, MenuLimitBoolean)
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "OSTA", SelectStr_OS, MenuLimitBoolean)
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "OSTP", SelectStr_OS, MenuLimitBoolean)

    'Set menu for Sclerosis
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "SCPF", SelectStr_Sclerosis, MenuLimitBoolean)

    'Set menu for Cysts
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "CYPF", SelectStr_PFCyst, MenuLimitBoolean)

    'Set menu for Chondrocalcinosis
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "CHON", SelectStr_MiscYN, MenuLimitBoolean)

    'Set menu for Joint Effusion
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "JE", SelectStr_MiscYN, MenuLimitBoolean)

    'Set menu for Ossification
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "OSQI", SelectStr_Ossification, MenuLimitBoolean)
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "OPTU", SelectStr_Ossification, MenuLimitBoolean)
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "OPTL", SelectStr_Ossification, MenuLimitBoolean)
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "OSLB", SelectStr_MiscYN, MenuLimitBoolean)

    'Set menu for PF OA
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "PFOA", SelectStr_MiscYN, MenuLimitBoolean)

    'Clear the error catching
    On Error GoTo 0

    Exit Function

ErrorHandler_Main1:
    Resume Next

End Function

'---LOADDROPDOWN_LLAT---'
Private Function LoadDropDown_LLat(ViewPrefix As String, MenuLimitBoolean As Boolean)
'Set dropdown menus for ComboBox controls on Left Lateral view subform

    Dim DummyBoolean As Boolean

    'Error catching
    On Error GoTo ErrorHandler_Main1

    'Set menu for PF KLG
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "PFKLG", SelectStr_PFKLG, MenuLimitBoolean)
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "PFOA", SelectStr_MiscYN, MenuLimitBoolean)

    'Set menu for PF JSN
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "PFJSN", SelectStr_JSN, MenuLimitBoolean)

    'Set menu for "FT" JSN from lateral view
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "FTJSM", SelectStr_JSN, MenuLimitBoolean)
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "FTJSL", SelectStr_JSN, MenuLimitBoolean)

    'Set menu for Osteophytes
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "OSFA", SelectStr_OS, MenuLimitBoolean)
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "OSFP", SelectStr_OS, MenuLimitBoolean)
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "OSPS", SelectStr_OS, MenuLimitBoolean)
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "OSPI", SelectStr_OS, MenuLimitBoolean)
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "OSTA", SelectStr_OS, MenuLimitBoolean)
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "OSTP", SelectStr_OS, MenuLimitBoolean)

    'Set menu for Sclerosis
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "SCPF", SelectStr_Sclerosis, MenuLimitBoolean)

    'Set menu for Cysts
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "CYPF", SelectStr_PFCyst, MenuLimitBoolean)

    'Set menu for Chondrocalcinosis
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "CHON", SelectStr_MiscYN, MenuLimitBoolean)

    'Set menu for Joint Effusion
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "JE", SelectStr_MiscYN, MenuLimitBoolean)

    'Set menu for Ossification
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "OSQI", SelectStr_Ossification, MenuLimitBoolean)
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "OPTU", SelectStr_Ossification, MenuLimitBoolean)
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "OPTL", SelectStr_Ossification, MenuLimitBoolean)
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "OSLB", SelectStr_MiscYN, MenuLimitBoolean)

    'Set menu for PF OA
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "PFOA", SelectStr_MiscYN, MenuLimitBoolean)

    'Clear the error catching
    On Error GoTo 0

    Exit Function

ErrorHandler_Main1:
    Resume Next

End Function

'---LOADAFTERUPDATE_PA---'
Private Function LoadAfterUpdate_PA(ViewPrefix As String)
'Set table SQL insertion for ComboBox controls on PA view subform

    Dim DummyBoolean As Boolean
    Dim TableName As String
    Dim FilterName1 As String
    Dim FilterValue1 As String

    'Default filter values
    TableName = "tblScores"
    FilterName1 = "READINGID"

    'Get READINGID from current record
    FilterValue1 = Me.Recordset.Fields("READINGID").Value

    'Error catching
    On Error GoTo ErrorHandler_Main1

    'Set update for TF KLG
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "TFKLG", TableName, FilterName1, FilterValue1)

    'Set update for TF JSN
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "TFJSM", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "TFJSL", TableName, FilterName1, FilterValue1)

    'Set update for Osteophytes
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "OSFM", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "OSFL", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "OSTM", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "OSTL", TableName, FilterName1, FilterValue1)

    'Set update for Sclerosis
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "SCFM", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "SCFL", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "SCTM", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "SCTL", TableName, FilterName1, FilterValue1)

    'Set update for Cysts
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "CYFM", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "CYFL", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "CYTM", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "CYTL", TableName, FilterName1, FilterValue1)

    'Set update for Attrition
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "ATTM", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "ATTL", TableName, FilterName1, FilterValue1)


    'Set update for Chondrocalcinosis
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "CHOM", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_PA", ViewPrefix, "CHOL", TableName, FilterName1, FilterValue1)

    'Clear the error catching
    On Error GoTo 0

    Exit Function

ErrorHandler_Main1:
    Resume Next

End Function

'---LOADAFTERUPDATE_RLAT---'
Private Function LoadAfterUpdate_RLat(ViewPrefix As String)
'Set table SQL insertion for ComboBox controls on Right Lateral view subform

    Dim DummyBoolean As Boolean
    Dim TableName As String
    Dim FilterName1 As String
    Dim FilterValue1 As String

    'Default filter values
    TableName = "tblScores"
    FilterName1 = "READINGID"

    'Get READINGID from current record
    FilterValue1 = Me.Recordset.Fields("READINGID").Value

    'Error catching
    On Error GoTo ErrorHandler_Main1

    'Set update for PF KLG
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "PFKLG", TableName, FilterName1, FilterValue1)

    'Set update for PF JSN
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "PFJSN", TableName, FilterName1, FilterValue1)

    'Set update for "FT" JSN from lateral view
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "FTJSM", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "FTJSL", TableName, FilterName1, FilterValue1)

    'Set update for Osteophytes
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "OSFA", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "OSFP", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "OSPS", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "OSPI", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "OSTA", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "OSTP", TableName, FilterName1, FilterValue1)

    'Set update for Sclerosis
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "SCPF", TableName, FilterName1, FilterValue1)

    'Set update for Cysts
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "CYPF", TableName, FilterName1, FilterValue1)

    'Set update for Chondrocalcinosis
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "CHON", TableName, FilterName1, FilterValue1)

    'Set update for Joint Effusion
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "JE", TableName, FilterName1, FilterValue1)

    'Set update for Ossification
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "OSQI", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "OPTU", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "OPTL", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "OSLB", TableName, FilterName1, FilterValue1)

    'Set update for PF OA
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "PFOA", TableName, FilterName1, FilterValue1)

    'Clear the error catching
    On Error GoTo 0

    Exit Function

ErrorHandler_Main1:
    Resume Next

End Function

'---LOADAFTERUPDATE_LLAT---'
Private Function LoadAfterUpdate_LLat(ViewPrefix As String)
'Set table SQL insertion ComboBox controls on Left Lateral view subform

    Dim DummyBoolean As Boolean
    Dim TableName As String
    Dim FilterName1 As String
    Dim FilterValue1 As String

    'Default filter values
    TableName = "tblScores"
    FilterName1 = "READINGID"

    'Get READINGID from current record
    FilterValue1 = Me.Recordset.Fields("READINGID").Value

    'Error catching
    On Error GoTo ErrorHandler_Main1

    'Set update for PF KLG
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "PFKLG", TableName, FilterName1, FilterValue1)

    'Set update for PF JSN
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "PFJSN", TableName, FilterName1, FilterValue1)

    'Set update for "FT" JSN from lateral view
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "FTJSM", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "FTJSL", TableName, FilterName1, FilterValue1)

    'Set update for Osteophytes
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "OSFA", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "OSFP", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "OSPS", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "OSPI", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "OSTA", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "OSTP", TableName, FilterName1, FilterValue1)

    'Set update for Sclerosis
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "SCPF", TableName, FilterName1, FilterValue1)

    'Set update for Cysts
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "CYPF", TableName, FilterName1, FilterValue1)

    'Set update for Chondrocalcinosis
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "CHON", TableName, FilterName1, FilterValue1)

    'Set update for Joint Effusion
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "JE", TableName, FilterName1, FilterValue1)

    'Set update for Ossification
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "OSQI", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "OPTU", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "OPTL", TableName, FilterName1, FilterValue1)
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "OSLB", TableName, FilterName1, FilterValue1)

    'Set update for PF OA
    DummyBoolean = SetComboUpdate_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "PFOA", TableName, FilterName1, FilterValue1)

    'Clear the error catching
    On Error GoTo 0

    Exit Function

ErrorHandler_Main1:
    Resume Next

End Function


'---COMMAND_LOAD_PA_CLICK---'
Private Sub Command_LOAD_PA_Click()

    Dim PATID_in As String
    Dim PATACC_in As String
    Dim NUMXR_var As Variant
    Dim NUMXR_in As Integer
    Dim DummyBoolean As Boolean

    On Error GoTo ErrorHandler1

    'Set subform focus
    Forms("Form_MOST_144_168").Controls("Subform_PA").SetFocus

    'Get current ID, Barcode, and number of images to display
    PATID_in = Me.Recordset("READINGID").Value
    PATACC_in = Me.Recordset("XBBARCDBU").Value
    NUMXR_var = Me.Recordset("NUMPA").Value

    If Len(Nz(NUMXR_var, "")) > 0 Then
        NUMXR_in = NUMXR_var
    Else
        NUMXR_in = 4
    End If

    'Load images in EFilm Workstation with API
    DummyBoolean = VBA_OpenStudy(PATID_in, PATACC_in, NUMXR_in)

    On Error GoTo 0
    Exit Sub

ErrorHandler1:
    Exit Sub

End Sub

'---COMMAND_LOAD_RL_CLICK---'
Private Sub Command_LOAD_RL_Click()

    Dim PATID_in As String
    Dim PATACC_in As String
    Dim NUMXR_var As Variant
    Dim NUMXR_in As Integer
    Dim DummyBoolean As Boolean

    On Error GoTo ErrorHandler1

    'Set subform focus
    Forms("Form_MOST_144_168").Controls("Subform_RLAT").SetFocus

    'Get current ID, Barcode, and number of images to display
    PATID_in = Me.Recordset("READINGID").Value
    PATACC_in = Me.Recordset("LXRBARCDBU").Value
    NUMXR_var = Me.Recordset("NUMRL").Value

    If Len(Nz(NUMXR_var, "")) > 0 Then
        NUMXR_in = NUMXR_var
    Else
        NUMXR_in = 4
    End If

    'Load images in EFilm Workstation with API
    DummyBoolean = VBA_OpenStudy(PATID_in, PATACC_in, NUMXR_in)

    On Error GoTo 0
    Exit Sub

ErrorHandler1:
    Exit Sub

End Sub

'---COMMAND_LOAD_LL_CLICK---'
Private Sub Command_LOAD_LL_Click()

    Dim PATID_in As String
    Dim PATACC_in As String
    Dim NUMXR_var As Variant
    Dim NUMXR_in As Integer
    Dim DummyBoolean As Boolean

    On Error GoTo ErrorHandler1

    'Set subform focus
    Forms("Form_MOST_144_168").Controls("Subform_LLAT").SetFocus

    'Get current ID, Barcode, and number of images to display
    PATID_in = Me.Recordset("READINGID").Value
    PATACC_in = Me.Recordset("LXLBARCDBU").Value
    NUMXR_var = Me.Recordset("NUMLL").Value

    If Len(Nz(NUMXR_var, "")) > 0 Then
        NUMXR_in = NUMXR_var
    Else
        NUMXR_in = 4
    End If

    'Load images in EFilm Workstation with API
    DummyBoolean = VBA_OpenStudy(PATID_in, PATACC_in, NUMXR_in)

    On Error GoTo 0
    Exit Sub

ErrorHandler1:
    Exit Sub

End Sub

'---COMMAND_UNLOCKV1V2_CLICK---'
Private Sub Command_UnlockV1V2_Click()
' Show/Unlock the V1 and V2 ComboBoxes to Reader for score updates

    Dim ControlName_RV12_PA() As String
    Dim ControlName_RV12_RLAT() As String
    Dim ControlName_RV12_LLAT() As String
    Dim DebugBoolean As Boolean
    Dim DummyBoolean As Boolean

    'Construct ComboBox Control Names from preloaded MOST variable root names
    ControlName_RV12_PA = Concat_Prefix("Combo_", MOST_RV12_XB_Vars)
    ControlName_RV12_RLAT = Concat_Prefix("Combo_", MOST_RV12_LXR_Vars)
    ControlName_RV12_LLAT = Concat_Prefix("Combo_", MOST_RV12_LXL_Vars)

    'Unlock the RV1 & RV2 PA variables on PA subform
    DummyBoolean = LockVars("Form_MOST_144_168", "Subform_PA", ControlName_RV12_PA, False)

    'Unlock the RIGHT knee RV1 & RV2 variables on RIGHT Lateral subform
    DummyBoolean = LockVars("Form_MOST_144_168", "Subform_RLAT", ControlName_RV12_RLAT, False)

    'Unlock the LEFT knee RV1 & RV2 variables on LEFT Lateral subform
    DummyBoolean = LockVars("Form_MOST_144_168", "Subform_LLAT", ControlName_RV12_LLAT, False)

    'Reload background color scheme for ComboBox Controls
    LoadColor

    'Error catching
    On Error GoTo ErrorHandler_Main1

    'Clear the error catching
    On Error GoTo 0

    Exit Sub

ErrorHandler_Main1:
  'Reload background color scheme for ComboBox Controls
  LoadColor
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

'---COMMAND_SUBMIT_CLICK---'
Private Sub Command_SUBMIT_Click()

    Dim CUR_DATE As String
    Dim CUR_TIME As String
    Dim CUR_READER As Variant

    Dim ComboNames_XB() As String
    Dim ComboNames_LXB() As String
    Dim nUnlocked_XB As Integer
    Dim nUnlocked_LXR As Integer
    Dim nUnlocked_LXL As Integer
    Dim nScored_XB As Integer
    Dim nScored_LXR As Integer
    Dim nScored_LXL As Integer

    Dim DummyBoolean As Boolean
    Dim MsgResponse As Integer
    Dim MsgExtra As Integer

    On Error GoTo ErrorHandler1

    'Get the current date and time
    CUR_DATE = Format(Now, "yyyymmdd")
    CUR_TIME = Format(Now, "hh:mm:ss")

    'Call for DirtySave
    DummyBoolean = DirtySave("Form_MOST_144_168")

    'Check for reader name
    CUR_READER = Me("Combo_HeaderREADER").Value
    If Len(Nz(CUR_READER, "")) < 1 Then
        MsgExtra = MsgBox("Please select reader name before signing the scoresheet.", vbOKOnly, "No Reader selected")
        Exit Sub
    End If

    'Check for empty scoring boxes
    ComboNames_XB = Concat_Prefix("Combo_", MOST_RV1234_XB_Vars)
    ComboNames_LXB = Concat_Prefix("Combo_", MOST_RV1234_LXB_Vars)

    nUnlocked_XB = CountUnlocked("Form_MOST_144_168", "Subform_PA", ComboNames_XB)
    nUnlocked_LXR = CountUnlocked("Form_MOST_144_168", "Subform_RLAT", ComboNames_LXB)
    nUnlocked_LXL = CountUnlocked("Form_MOST_144_168", "Subform_LLAT", ComboNames_LXB)

    nScored_XB = CountScored("Form_MOST_144_168", "Subform_PA", ComboNames_XB)
    nScored_LXR = CountScored("Form_MOST_144_168", "Subform_RLAT", ComboNames_LXB)
    nScored_LXL = CountScored("Form_MOST_144_168", "Subform_LLAT", ComboNames_LXB)

    'for debugging
    'MsgResponse = MsgBox(CStr(nUnlocked_XB) & " " & CStr(nUnlocked_LXR) & " " & CStr(nUnlocked_LXL), vbYesNo + vbCritical + vbDefaultButton2, "Confirm")
    'MsgResponse = MsgBox(CStr(nScored_XB) & " " & CStr(nScored_LXR) & " " & CStr(nScored_LXL), vbYesNo + vbCritical + vbDefaultButton2, "Confirm")

    'Sign the record


    'Exit the Sub
    On Error GoTo 0
    Exit Sub

ErrorHandler1:
    Exit Sub

End Sub
