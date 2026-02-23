VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' --- Event Hook ---
Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    ' Delegate logic to the standard module to keep ThisWorkbook clean
    Module_EnumSelector.TryLaunchEnumSelector Target
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Clear the cache to release memory before Excel closes
    ' We silently clear the dictionary without popping the MsgBox
    Module_EnumSelector.SilentRefreshCache
End Sub
