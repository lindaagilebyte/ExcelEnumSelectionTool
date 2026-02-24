VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_EnumSelect 
   Caption         =   "Select Value"
   ClientHeight    =   4000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   StartUpPosition =   1  'CenterOwner
   Begin {A22D307D-8973-41E7-862D-0FCCE8A28E3F} lblHeader 
      Caption         =   "Please select a value:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin {8BD21D13-EC42-11CE-9E0D-00AA006002F3} lstEnums 
      Height          =   2640
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4455
   End
   Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnRefresh 
      Caption         =   "Refresh Cache"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   4455
   End
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
    Me.Caption = "選擇數值: " & key
    Me.lblHeader.Caption = "請為 " & key & " 選擇一個數值:"
    Me.btnRefresh.Caption = "重新整理快取"
    
    Me.lstEnums.Clear
    
    Dim item As Variant
    For Each item In data
        Me.lstEnums.AddItem item
    Next item
End Sub

' --- Event Handlers ---
Private Sub lstEnums_Click()
    ' User just clicked an item - do nothing until Confirm is pressed
End Sub

Private Sub btnConfirm_Click()
    ' Ensure an item is actually selected
    If IsNull(Me.lstEnums.Value) Then
        MsgBox "請先選擇一個選項。", vbExclamation, "提示"
        Exit Sub
    End If

    ' Save State for Undo
    Set Module_EnumSelector.pUndoSheet = ActiveSheet
    Set Module_EnumSelector.pUndoCell = ActiveCell
    Module_EnumSelector.pUndoValue = ActiveCell.Value

    ' Write to Active Cell
    ActiveCell.Value = Me.lstEnums.Value
    
    ' Register Custom Undo
    Application.OnUndo "復原列舉選擇", "Module_EnumSelector.UndoEnumSelection"
    
    ' Close Form
    Unload Me
End Sub

Private Sub btnCancel_Click()
    ' Cancel operation and close form
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    ' Default setup happens in InitializeWithData
End Sub

Private Sub btnRefresh_Click()
    ' Delegate to Module to clear cache
    Module_EnumSelector.RefreshCache
    
    ' Close Form
    Unload Me
    MsgBox "快取已清除。請再次點擊儲存格以重新載入資料。", vbInformation
End Sub
