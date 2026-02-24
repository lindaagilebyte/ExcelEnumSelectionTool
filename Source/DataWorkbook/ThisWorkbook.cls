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
Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    ' Delegate logic to the standard module
    ' If the module successfully launches the Form, we Cancel the native double-click edit mode
    If Module_EnumSelector.TryLaunchEnumSelector(Target) Then
        Cancel = True
    End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Clear the cache to release memory before Excel closes
    ' We silently clear the dictionary without popping the MsgBox
    Module_EnumSelector.SilentRefreshCache
End Sub
