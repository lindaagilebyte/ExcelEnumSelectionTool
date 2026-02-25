Attribute VB_Name = "Module_EnumSelector"
Option Explicit

' --- Constants ---
Public Const CONST_DEBUG_MODE As Boolean = True ' Set to True to output log to Immediate Window
Private Const REF_FILE_NAME As String = "列舉定義(企劃用).xlsx"
Private Const DATA_SUB_HEADER As String = "定義(巨集顯示)"

' --- Global Cache ---
' Dictionary: Key = EnumName (String), Value = Variant Array of Strings
Private pEnumCache As Object

' --- Undo State ---
Public pUndoSheet As Worksheet
Public pUndoCell As Range
Public pUndoValue As Variant

' --- Entry Point ---
' Called by Workbook_SheetBeforeDoubleClick in ThisWorkbook
Public Function TryLaunchEnumSelector(Target As Range) As Boolean
    On Error GoTo ErrorHandler
    
    TryLaunchEnumSelector = False ' Default
    
    ' 1. Row Validation: Only activate for Row >= 5
    If Target.Row < 5 Then Exit Function
    If Target.Cells.Count > 1 Then Exit Function ' Multi-select ignored
    
    ' 2. Column Validation: Check Row 3 for Enum Key (Header)
    Dim enumKey As String
    ' Note: We assume the "Enum Key" is always in Row 3 of the data sheet.
    enumKey = Trim(CStr(Target.Worksheet.Cells(3, Target.Column).Value))
    
    If Len(enumKey) = 0 Then Exit Function
    
    ' 3. Load Definitions (with Caching)
    Dim enumList As Variant
    enumList = GetEnumList(enumKey)
    
    ' 4. Launch UserForm if data found
    If IsArray(enumList) Then
        If UBound(enumList) >= LBound(enumList) Then
            ' Pass data to Form
            Form_EnumSelect.InitializeWithData enumKey, enumList
            TryLaunchEnumSelector = True
            Form_EnumSelect.Show vbModal
        Else
            MsgBox "找不到 [" & enumKey & "] 的資料定義，請檢查列舉參考檔。", vbExclamation, "列舉定義缺失"
        End If
    End If
    
    Exit Function

ErrorHandler:
    If CONST_DEBUG_MODE Then Debug.Print "[DEBUG] Error in TryLaunchEnumSelector: " & Err.Description
    TryLaunchEnumSelector = False
End Function

' --- Cache Management ---
Private Function GetEnumList(key As String) As Variant
    ' Initialize Cache if needed (Cold Start)
    If pEnumCache Is Nothing Then
        Set pEnumCache = CreateObject("Scripting.Dictionary")
        ScanReferenceFile GetRequiredEnumKeys()
    End If
    
    ' Warm Start Look up
    If pEnumCache.Exists(key) Then
        GetEnumList = pEnumCache(key)
    Else
        GetEnumList = Null
    End If
End Function

Public Sub RefreshCache()
    Set pEnumCache = Nothing
    MsgBox "快取已清除。下次點擊將重新從參考檔載入。", vbInformation, "Enum Selector"
End Sub

Public Sub SilentRefreshCache()
    ' Called silently during Workbook_BeforeClose
    Set pEnumCache = Nothing
End Sub

' --- Undo Management ---
Public Sub UndoEnumSelection()
    On Error Resume Next
    If Not pUndoCell Is Nothing Then
        ' Ensure we are on the correct sheet before restoring
        If pUndoSheet.Name = ActiveSheet.Name Then
            pUndoCell.Value = pUndoValue
        End If
    End If
    On Error GoTo 0
End Sub

' --- Reference File Scanning ---
Private Function GetRequiredEnumKeys() As Object
    Dim requiredKeys As Object
    Set requiredKeys = CreateObject("Scripting.Dictionary")
    
    Dim ws As Worksheet
    Dim rng As Range
    Dim c As Range
    Dim val As String
    
    ' Scan all sheets in the current Data File (ThisWorkbook)
    For Each ws In ThisWorkbook.Worksheets
        ' Ignore utility or description sheets starting with #
        If Left(ws.Name, 1) <> "#" Then
            On Error Resume Next
            Set rng = ws.UsedRange
            On Error GoTo 0
            
            If Not rng Is Nothing Then
                ' Check if Row 3 exists within the UsedRange
                If rng.Row <= 3 And (rng.Row + rng.Rows.Count - 1) >= 3 Then
                    ' Scan across Row 3 for valid keys
                    Dim colIdx As Long
                    Dim startCol As Long, endCol As Long
                    startCol = rng.Column
                    endCol = rng.Column + rng.Columns.Count - 1
                    
                    For colIdx = startCol To endCol
                        val = Trim(CStr(ws.Cells(3, colIdx).Value))
                        If Len(val) > 0 Then
                            If Not requiredKeys.Exists(val) Then
                                requiredKeys.Add val, True
                            End If
                        End If
                    Next colIdx
                End If
            End If
        End If
    Next ws
    
    Set GetRequiredEnumKeys = requiredKeys
End Function

Private Sub ScanReferenceFile(requiredKeys As Object)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 1. Resolve Path
    Dim wbPath As String, refPath As String
    wbPath = ThisWorkbook.Path
    
    ' Strategy 1: Sibling
    refPath = fso.BuildPath(wbPath, REF_FILE_NAME)
    
    ' Strategy 2: Parent (SVN style - typical for "Reference" folder sibling to "Form" folder)
    If Not fso.FileExists(refPath) Then
        refPath = fso.BuildPath(fso.GetParentFolderName(wbPath), REF_FILE_NAME)
    End If
    
    ' Strategy 3: Check "reference" subfolder if we are in root (Development/Agent context)
    If Not fso.FileExists(refPath) Then
         Dim devPath As String
         devPath = fso.BuildPath(wbPath, "reference\" & REF_FILE_NAME)
         If fso.FileExists(devPath) Then refPath = devPath
    End If
    
    If CONST_DEBUG_MODE Then Debug.Print "[DEBUG] Final target path for Reference: " & refPath
    
    ' Strategy 4: Manual Pick
    If Not fso.FileExists(refPath) Then
        If MsgBox("Reference file not found: " & REF_FILE_NAME & vbCrLf & "Browse to select it?", vbQuestion + vbYesNo) = vbYes Then
            Dim fd As FileDialog
            Set fd = Application.FileDialog(msoFileDialogFilePicker)
            fd.Title = "Select " & REF_FILE_NAME
            If fd.Show = -1 Then
                refPath = fd.SelectedItems(1)
            Else
                Exit Sub ' User cancelled
            End If
        Else
            Exit Sub
        End If
    End If
    
    ' 2. Open Workbook Read-Only
    Dim sourceWb As Workbook
    Dim screenUpdateState As Boolean
    
    screenUpdateState = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    Dim tempPath As String
    ' Generate a unique filename using a GUID to prevent read-locking collisions from zombie tasks
    Dim tempGuid As String
    tempGuid = Replace(Mid(CreateObject("Scriptlet.TypeLib").Guid, 2, 36), "-", "")
    tempPath = Environ("TEMP") & "\" & tempGuid & "_" & REF_FILE_NAME
    
    On Error Resume Next
    If fso.FileExists(tempPath) Then fso.DeleteFile tempPath, True
    fso.CopyFile refPath, tempPath, True
    On Error GoTo 0
    
    ' SVN often applies strict read-only locks. We just copy it to TEMP and open that copy safely.
    On Error Resume Next
    Set sourceWb = Workbooks.Open(Filename:=tempPath, ReadOnly:=True, UpdateLinks:=False, IgnoreReadOnlyRecommended:=True)
    On Error GoTo 0
    
    If sourceWb Is Nothing Then
        Application.ScreenUpdating = screenUpdateState
        MsgBox "無法開啟列舉參考檔。", vbCritical
        Exit Sub
    End If
    
    ' 3. Scan Sheets
    Dim ws As Worksheet
    Dim totalSheets As Long, currentSheet As Long
    totalSheets = sourceWb.Worksheets.Count
    currentSheet = 1
    
    For Each ws In sourceWb.Worksheets
        Application.StatusBar = "正在讀取列舉定義... (Sheet " & currentSheet & " of " & totalSheets & ")"
        ScanWorksheet ws, requiredKeys
        currentSheet = currentSheet + 1
    Next ws
    
    ' 4. Cleanup
    Application.StatusBar = False
    sourceWb.Close SaveChanges:=False
    
    ' Delete temp file to keep system clean
    On Error Resume Next
    fso.DeleteFile tempPath, True
    On Error GoTo 0
    
    Application.ScreenUpdating = screenUpdateState
    
    
    If CONST_DEBUG_MODE Then Debug.Print "[DEBUG] Cache built. Total Enums cached: " & pEnumCache.Count
End Sub

Private Sub ScanWorksheet(ws As Worksheet, requiredKeys As Object)
    On Error Resume Next
    Dim rng As Range
    Set rng = ws.UsedRange
    If rng Is Nothing Then Exit Sub
    On Error GoTo 0
    
    Dim c As Range
    Dim firstAddress As String
    
    ' Search for the sub-header string
    Set c = rng.Find(What:=DATA_SUB_HEADER, LookIn:=xlValues, LookAt:=xlWhole)
    If Not c Is Nothing Then
        firstAddress = c.Address
        Do
            ' Found a definition block.
            ' Assumption: The "Enum Key" is located at (Row-1, Col-1) relative to this sub-header.
            ' Example: Key at A10, Sub-Header at B11.
            
            If c.Row > 1 And c.Column > 1 Then
                Dim keyCell As Range
                Set keyCell = ws.Cells(c.Row - 1, c.Column - 1)
                
                Dim keyName As String
                keyName = Trim(CStr(keyCell.Value))
                
                ' Parse Data if Key is valid, needed by Data file, and not already cached
                If Len(keyName) > 0 Then
                     If requiredKeys.Exists(keyName) Then
                         If Not pEnumCache.Exists(keyName) Then
                            Dim items As Variant
                            items = ExtractColumnData(ws, c.Row + 1, c.Column)
                            
                            ' Only add if we found items
                            If UBound(items) >= 0 Then
                                pEnumCache.Add keyName, items
                            End If
                        End If
                    End If
                End If
            End If
            
            Set c = rng.FindNext(c)
        Loop While Not c Is Nothing And c.Address <> firstAddress
    End If
End Sub

' Extracts a vertical list starting from (startRow, colIndex) downwards until empty cell
' Returns a Variant Array (0-based)
Private Function ExtractColumnData(ws As Worksheet, startRow As Long, colIndex As Long) As Variant
    Dim dataList() As String
    ReDim dataList(0 To 99) ' Initial buffer
    Dim count As Long
    count = 0
    
    Dim r As Long
    r = startRow
    
    Do
        Dim val As String
        val = Trim(CStr(ws.Cells(r, colIndex).Value))
        
        ' Stop at empty cell
        If Len(val) = 0 Then Exit Do
        
        ' Resize buffer if needed
        If count > UBound(dataList) Then
            ReDim Preserve dataList(0 To UBound(dataList) + 100)
        End If
        
        dataList(count) = val
        count = count + 1
        r = r + 1
        
        ' Safety break to prevent infinite loops
        If r > startRow + 5000 Then Exit Do
    Loop
    
    If count > 0 Then
        ' Trim to exact size
        ReDim Preserve dataList(0 To count - 1)
        ExtractColumnData = dataList
    Else
        ExtractColumnData = Array()
    End If
End Function
