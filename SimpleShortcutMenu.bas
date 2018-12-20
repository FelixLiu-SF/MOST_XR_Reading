Attribute VB_Name = "SimpleShortcutMenu"
Option Compare Database

 
Sub CreateSimpleShortcutMenu()

    Dim cmbShortcutMenu As Office.CommandBar

    ' Create a shortcut menu named "SimpleShortcutMenu.
    Set cmbShortcutMenu = CommandBars.Add("SimpleShortcutMenu", msoBarPopup, False, True)

    ' Add the Cut command.
    cmbShortcutMenu.Controls.Add Type:=msoControlButton, Id:=21

    ' Add the Copy command.
    cmbShortcutMenu.Controls.Add Type:=msoControlButton, Id:=19

    ' Add the Paste command.
    cmbShortcutMenu.Controls.Add Type:=msoControlButton, Id:=22

    Set cmbShortcutMenu = Nothing
     
End Sub
