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


Private Sub SetComboSelection_RV1234(FormName As String, SubFormControlName As String, VarNameRoot As String, SelectionStr As String, MenuLimitBoolean)

    Dim DummyBoolean As Boolean
    Dim VisitArray(4) As String
    Dim ControlName As String
    Dim OnFocusStr As String
    Dim Index As Integer
    
    'Define default variables
    VisitArray(0) = "RV1"
    VisitArray(1) = "RV1"
    VisitArray(2) = "RV1"
    VisitArray(3) = "RV1"
    
    'Loop through visits
    Index = 0
    For Index = 0 To 4
    
        'Construct ComboBox Control name
        ControlName = "Combo_" & VarNameRoot
        
        'Construct ComboBox selection string
        OnFocusStr = Make_ControlUpdate_Func(FormName, SubFormControlName, ControlName, SelectionStr)
        
        'Set the selection string to the OnFocus property of the ComboBox
        DummyBoolean = Control_Edit_OnFocus(FormName, SubFormControlName, ControlName, OnFocusStr)
        
        'Set the ComboBox dropdown menu display properties
        DummmyBoolean = Control_Edit_Binding(FormName, SubFormControlName, ControlName, 4, 2, "0; 0; 0.5 in; 2 in", 3, MenuLimitBoolean)

End Sub
