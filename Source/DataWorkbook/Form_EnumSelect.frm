VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_EnumSelect 
   Caption         =   "Select Value"
   ClientHeight    =   4000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "Form_EnumSelect.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form_EnumSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' --- Private Variables ---
Private pTargetKey As String

' --- Public API ---
Public Sub InitializeWithData(key As String, data As Variant)
    pTargetKey = key
    Me.Caption = "Select: " & key
    
    Me.lstEnums.Clear
    
    Dim item As Variant
    For Each item In data
        Me.lstEnums.AddItem item
    Next item
    
    ' Resize form based on content (Optional, stick to fixed size for now)
End Sub

' --- Event Handlers ---

Private Sub lstEnums_Click()
    ' Write to Active Cell
    ActiveCell.Value = Me.lstEnums.Value
    
    ' Close Form
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    ' Default setup
    Me.lblHeader.Caption = "Please select a value:"
End Sub

Private Sub btnRefresh_Click()
    ' Delegate to Module to clear cache
    Module_EnumSelector.RefreshCache
    
    ' Reload (attempt to re-trigger or just close)
    Unload Me
    MsgBox "Cache refreshed. Please click the cell again.", vbInformation
End Sub

' --- Control Declarations (Simulated for Import) ---
' These would normally be in the .frx or designer
' ListBox: lstEnums
' Label: lblHeader
' CommandButton: btnRefresh
