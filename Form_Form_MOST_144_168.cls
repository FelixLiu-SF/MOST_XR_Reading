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

Private Sub Form_Load()
' Load various modules when this scoresheet is opened

    Dim DummyBoolean As Boolean

    'Error catching
    On Error GoTo ErrorHandler_Main1

    'Load the dataset objects for combo box dropdown menus
    LoadDAO

    'Load the EFilm Automation object to link with EFilm Workstation
    LoadEFilmAuto

    'Load MOST global variable names and SQL selection strings
    DummyBoolean = MOST_Load_SelectStr
    DummyBoolean = MOST_Load_VariableNameArrays

    'Load default view for ComboBox Controls
    LoadView

    'Set all ComboBox OnLostFocus properties to color coding
    SetComboLostFocus

    'Set all ComboBox OnFocus properties to dropdown menus
    DummyBoolean = LoadDropDown

    'Load initial background color scheme for ComboBox Controls
    LoadColor

    'Clear the error catching
    On Error GoTo 0

    Exit Sub

ErrorHandler_Main1:
    Resume Next

End Sub

Private Sub LoadView()
' Show/Hide/Lock/Unlock the ComboBoxes for default view on all subforms

    Dim ControlName_RV12_PA() As String
    Dim ControlName_RV12_RLAT() As String
    Dim ControlName_RV12_LLAT() As String
    Dim ControlName_RV34_RLAT() As String
    Dim ControlName_RV34_LLAT() As String
    Dim DummyBoolean As Boolean

    'Error catching
    On Error GoTo ErrorHandler_Main1

    'Construct ComboBox Control Names from preloaded MOST variable root names
    ControlName_RV12_PA = Concat_Prefix("Combo_", MOST_RV12_XB_Vars)
    ControlName_RV12_RLAT = Concat_Prefix("Combo_", MOST_RV12_LXR_Vars)
    ControlName_RV12_LLAT = Concat_Prefix("Combo_", MOST_RV12_LXL_Vars)
    ControlName_RV34_RLAT = Concat_Prefix("Combo_", MOST_RV34_LXR_Vars)
    ControlName_RV34_LLAT = Concat_Prefix("Combo_", MOST_RV34_LXL_Vars)

    'Lock the RV1 & RV2 PA variables on PA subform
    DummyBoolean = LockVars("Form_MOST_144_168", "Subform_PA", ControlName_RV12_PA, True)

    'Lock the RIGHT knee RV1 & RV2 variables on RIGHT Lateral subform
    DummyBoolean = LockVars("Form_MOST_144_168", "Subform_RLAT", ControlName_RV12_RLAT, True)

    'Lock the LEFT knee RV1 & RV2 variables on LEFT Lateral subform
    DummyBoolean = LockVars("Form_MOST_144_168", "Subform_LLAT", ControlName_RV12_LLAT, True)

    'Lock & Hide all of the LEFT knee variables on RIGHT Lateral subform
    DummyBoolean = LockVars("Form_MOST_144_168", "Subform_RLAT", ControlName_RV12_LLAT, True)
    DummyBoolean = HideVars("Form_MOST_144_168", "Subform_RLAT", ControlName_RV12_LLAT, True)

    DummyBoolean = LockVars("Form_MOST_144_168", "Subform_RLAT", ControlName_RV34_LLAT, True)
    DummyBoolean = HideVars("Form_MOST_144_168", "Subform_RLAT", ControlName_RV34_LLAT, True)

    'Lock & Hide all of the RIGHT knee variables on LEFT Lateral subform
    DummyBoolean = LockVars("Form_MOST_144_168", "Subform_LLAT", ControlName_RV12_RLAT, True)
    DummyBoolean = HideVars("Form_MOST_144_168", "Subform_LLAT", ControlName_RV12_RLAT, True)

    DummyBoolean = LockVars("Form_MOST_144_168", "Subform_LLAT", ControlName_RV34_RLAT, True)
    DummyBoolean = HideVars("Form_MOST_144_168", "Subform_LLAT", ControlName_RV34_RLAT, True)

    'Clear the error catching
    On Error GoTo 0

    Exit Sub

ErrorHandler_Main1:
    Resume Next

End Sub


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

Private Function LoadDropDown()
'Set dropdown menus for all ComboBox controls on all subforms

    Dim DummyBoolean As Boolean

    'Error catching
    On Error GoTo ErrorHandler_Main1

    'Set menus on PA view subform & limit selections to menu list
    DummyBoolean = LoadDropDown_PA("XR", True)
    DummyBoolean = LoadDropDown_PA("XL", True)

    'Set menus on Right Lateral view subform & limit selections to menu list
    DummyBoolean = LoadDropDown_RLat("LXR", True)

    'Set menus on Left Lateral view subform & limit selections to menu list
    DummyBoolean = LoadDropDown_LLat("LXL", True)

    'Clear the error catching
    On Error GoTo 0

    Exit Sub

ErrorHandler_Main1:
    Resume Next

End Function

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

    Exit Sub

ErrorHandler_Main1:
    Resume Next

End Function

Private Function LoadDropDown_RLat(ViewPrefix As String, MenuLimitBoolean As Boolean)
'Set dropdown menus for ComboBox controls on Right Lateral view subform

    'Error catching
    On Error GoTo ErrorHandler_Main1

    'Set menu for PF KLG
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "PFKLG", SelectStr_PFKLG, MenuLimitBoolean)

    'Set menu for PF JSN
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_RLAT", ViewPrefix, "PFJSMN", SelectStr_JSN, MenuLimitBoolean)

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


    'Clear the error catching
    On Error GoTo 0

    Exit Sub

ErrorHandler_Main1:
    Resume Next

End Function

Private Function LoadDropDown_LLat(ViewPrefix As String, MenuLimitBoolean As Boolean)
'Set dropdown menus for ComboBox controls on Left Lateral view subform

    'Error catching
    On Error GoTo ErrorHandler_Main1

    'Set menu for PF KLG
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "PFKLG", SelectStr_PFKLG, MenuLimitBoolean)

    'Set menu for PF JSN
    DummyBoolean = SetComboSelection_RV1234("Form_MOST_144_168", "Subform_LLAT", ViewPrefix, "PFJSMN", SelectStr_JSN, MenuLimitBoolean)

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


    'Clear the error catching
    On Error GoTo 0

    Exit Sub

ErrorHandler_Main1:
    Resume Next

End Function


Private Function SetComboSelection_RV1234(FormName As String, SubFormControlName As String, ViewPrefix As String, VarNameRoot As String, SelectionStr As String, MenuLimitBoolean As Boolean)

    Dim DummyBoolean As Boolean
    Dim VisitArray(4) As String
    Dim ControlName As String
    Dim OnFocusStr As String
    Dim Index As Integer

    'Define default variables
    VisitArray(0) = "RV1"
    VisitArray(1) = "RV2"
    VisitArray(2) = "RV3"
    VisitArray(3) = "RV4"

    'Loop through visits
    Index = 0
    For Index = 0 To 4

        'Construct ComboBox Control name
        ControlName = "Combo_" & ViewPrefix & VarNameRoot

        'Construct ComboBox selection string
        OnFocusStr = Make_ControlUpdate_Func(FormName, SubFormControlName, ControlName, SelectionStr)

        'Set the selection string to the OnFocus property of the ComboBox
        DummyBoolean = Control_Edit_OnFocus(FormName, SubFormControlName, ControlName, OnFocusStr)

        'Set the ComboBox dropdown menu display properties
        DummyBoolean = Control_Edit_Binding(FormName, SubFormControlName, ControlName, 4, 2, "0; 0; 0.5 in; 2 in", 3, MenuLimitBoolean)

End Function
