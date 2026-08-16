Attribute VB_Name = "modExport"
Option Explicit

'=====================================================================================
' modExport - flat, values-only table for Power Automate / Power BI / Power Query
'
' *** เวอร์ชันแก้บั๊กหัวตารางเลื่อน (header shift) ***
'
' ปัญหาที่พบในไฟล์ V2
' -------------------
'   BuildHeaders เดิมประกาศชื่อคอลัมน์ ASCII ไว้ 30 ชื่อ โดยมี "OnhandAfter"
'   แทรกอยู่ตำแหน่งที่ 10 แต่ตารางต้นทาง tblMA_RO2RO3 มีเพียง 29 คอลัมน์
'   และไม่มีคอลัมน์ OnhandAfter
'
'   ผลคือหัวตารางตั้งแต่คอลัมน์ J ถึง AC "เลื่อนไป 1 ช่อง" เทียบกับข้อมูลจริง
'   ตรวจสอบได้จากไฟล์ V2 โดยตรง:
'
'       คอลัมน์  หัวที่เขียนไว้        ข้อมูลที่อยู่จริง
'       -------  -------------------  ----------------------------
'       J        OnhandAfter          จัดสรร            (Alloc)
'       K        Alloc                เบิกจาก Hub       (FromHub)
'       L        FromHub              จัดสรรเพิ่ม        (ExtraAlloc)
'       ...
'       T        PrevBefore           Prev_Received  <-- ผิดคนละคอลัมน์
'       U        PrevReceived         Expected_Before
'       W        NetUsed              Days_Period    <-- ค่า 7 คือจำนวนวัน ไม่ใช่ยอดใช้
'       AA       SuggestNextAlloc     Status
'       AB       Status               Risk_Flag
'       AC       RiskFlag             Data_Set       <-- ค่า "MA"/"Optimize"
'
'   ถ้าดึง tblPAExport เข้า Power Automate โดยไม่แก้ การ์ดจะแสดง
'   Prev_Received ในช่อง Prev_Before, แสดง Risk_Flag ในช่อง Status
'   และ Data_Set จะไปโผล่ในช่อง RiskFlag — ผิดทั้งใบโดยไม่มี error ให้เห็น
'
' การแก้
' ------
'   1. ตัด "OnhandAfter" ออกจากลิสต์ ให้เหลือ 29 ชื่อ ตรงกับตารางต้นทาง
'   2. เพิ่ม VerifyExportHeaders() ไว้ตรวจว่าจำนวนคอลัมน์ต้นทางยังตรงกับลิสต์
'      ถ้าไม่ตรง (เช่นมีคนเพิ่มคอลัมน์ในอนาคต) จะเตือนทันที แทนที่จะเงียบแล้วเพี้ยน
'
' วิธีติดตั้ง
' ---------
'   เปิด VBA Editor (Alt+F11) -> คลิกขวาที่ modExport -> Remove modExport
'   -> File -> Import File... -> เลือกไฟล์นี้ -> รัน BuildExportTable_Click อีกครั้ง
'=====================================================================================

Public Const SH_EXPORT  As String = "PA_Export"
Public Const TBL_EXPORT As String = "tblPAExport"

Private Const EXTRA_COLS As Long = 2      ' CutoffDate + ExportedAt


Public Sub BuildExportTable_Click()
    Dim n As Long

    If Not VerifyExportHeaders(True) Then Exit Sub

    n = BuildExportTable()
    MsgBox "PA_Export rebuilt." & vbCrLf & vbCrLf & _
           "Rows  : " & n & vbCrLf & _
           "Table : " & TBL_EXPORT & vbCrLf & _
           "Sheet : " & SH_EXPORT & vbCrLf & vbCrLf & _
           "Values only - safe for Power Automate.", vbInformation, "Export"
End Sub


'-------------------------------------------------------------------------------------
' ตรวจว่าจำนวนคอลัมน์ของตารางต้นทางยังตรงกับลิสต์ชื่อ ASCII หรือไม่
' คืนค่า True เมื่อปลอดภัยที่จะ export
'-------------------------------------------------------------------------------------
Public Function VerifyExportHeaders(Optional ByVal showUI As Boolean = False) As Boolean
    Dim lo   As ListObject
    Dim def  As Variant
    Dim c    As Long
    Dim bad  As String

    Set lo = SourceTable()
    If lo Is Nothing Then
        If showUI Then MsgBox "ไม่พบตารางต้นทาง " & modRefreshMA.TBL_MAIN, vbCritical, "Export"
        Exit Function
    End If

    def = HeaderNames()

    If lo.ListColumns.Count <> UBound(def) + 1 Then
        If showUI Then
            MsgBox "จำนวนคอลัมน์ไม่ตรงกัน!" & vbCrLf & vbCrLf & _
                   "ตารางต้นทาง : " & lo.ListColumns.Count & " คอลัมน์" & vbCrLf & _
                   "ลิสต์ชื่อ ASCII : " & UBound(def) + 1 & " ชื่อ" & vbCrLf & vbCrLf & _
                   "ถ้า export ต่อ หัวตารางจะเลื่อนไม่ตรงกับข้อมูล" & vbCrLf & _
                   "กรุณาแก้ฟังก์ชัน HeaderNames() ใน modExport ให้ตรงกับตารางก่อน", _
                   vbCritical, "Export - header mismatch"
        End If
        modRefreshMA.SysLog "Export ABORT - column count mismatch: source=" & _
                            lo.ListColumns.Count & " names=" & UBound(def) + 1
        Exit Function
    End If

    ' เตือนแบบไม่บล็อก ถ้าคอลัมน์ ASCII ตัวแรก ๆ ไม่ได้เรียงอย่างที่คาด
    For c = 1 To lo.ListColumns.Count
        If LenB(Trim$(lo.ListColumns(c).Name)) = 0 Then
            bad = bad & "  - คอลัมน์ที่ " & c & " ไม่มีชื่อ" & vbCrLf
        End If
    Next c
    If LenB(bad) > 0 Then modRefreshMA.SysLog "Export warn:" & vbCrLf & bad

    VerifyExportHeaders = True
End Function


'-------------------------------------------------------------------------------------
' Rebuilds PA_Export from tblMA_RO2RO3. Returns the number of data rows written.
'-------------------------------------------------------------------------------------
Public Function BuildExportTable() As Long
    Dim wsOut   As Worksheet
    Dim loSrc   As ListObject
    Dim src     As Variant
    Dim out()   As Variant
    Dim hdr()   As String
    Dim cutoff  As Object
    Dim nRows   As Long
    Dim nSrcCol As Long
    Dim nOutCol As Long
    Dim r       As Long
    Dim c       As Long
    Dim stamp   As String
    Dim per     As String
    Dim oldAF   As Boolean

    Set loSrc = SourceTable()
    If loSrc Is Nothing Then Exit Function

    nSrcCol = loSrc.ListColumns.Count
    nOutCol = nSrcCol + EXTRA_COLS
    nRows = loSrc.ListRows.Count

    oldAF = modRefreshMA.GetAutoFillLists()
    modRefreshMA.SetAutoFillLists False

    Set wsOut = EnsureExportSheet()
    hdr = BuildHeaders(loSrc, nOutCol)
    Set cutoff = CutoffDict()
    stamp = Format$(Now, "yyyy-mm-dd hh:nn:ss")

    ' header row
    For c = 1 To nOutCol
        wsOut.Cells(1, c).Value = hdr(c)
    Next c

    If nRows = 0 Then
        ResizeExportTable wsOut, 1, nOutCol
        modRefreshMA.SetAutoFillLists oldAF
        Exit Function
    End If

    ' read every value once - Value2 gives raw numbers, never a formula
    If nRows = 1 Then
        ReDim src(1 To 1, 1 To nSrcCol)
        For c = 1 To nSrcCol
            src(1, c) = loSrc.DataBodyRange.Cells(1, c).Value2
        Next c
    Else
        src = loSrc.DataBodyRange.Value2
    End If

    ReDim out(1 To nRows, 1 To nOutCol)
    For r = 1 To nRows
        For c = 1 To nSrcCol
            out(r, c) = CleanCell(src(r, c))
        Next c
        per = Trim$(modRefreshMA.SafeTxt(src(r, 1)))
        out(r, nSrcCol + 1) = IsoDate(cutoff, per)
        out(r, nSrcCol + 2) = stamp
    Next r

    wsOut.Range("A2").Resize(nRows, nOutCol).ClearContents
    wsOut.Range("A2").Resize(nRows, nOutCol).NumberFormat = "General"
    wsOut.Range("A2").Resize(nRows, nOutCol).Value = out

    ClearBelow wsOut, nRows + 1, nOutCol

    ResizeExportTable wsOut, nRows + 1, nOutCol
    wsOut.Columns.AutoFit

    modRefreshMA.SetAutoFillLists oldAF
    modRefreshMA.SysLog "PA_Export rebuilt - " & nRows & " rows"
    BuildExportTable = nRows
End Function


'=====================================================================================
' helpers
'=====================================================================================

Private Function SourceTable() As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(modRefreshMA.SH_OVER)
    If Not ws Is Nothing Then Set lo = ws.ListObjects(modRefreshMA.TBL_MAIN)
    On Error GoTo 0

    Set SourceTable = lo
End Function


'-------------------------------------------------------------------------------------
' ชื่อคอลัมน์ ASCII 29 ชื่อ เรียงตรงกับ tblMA_RO2RO3 คอลัมน์ A..AC
'
' ***  จุดที่แก้บั๊ก  ***
' ของเดิมมี "OnhandAfter" แทรกอยู่หลัง "OnhandBefore" รวมเป็น 30 ชื่อ
' ทั้งที่ตารางต้นทางไม่มีคอลัมน์นั้น ทำให้หัวตารางเลื่อนไป 1 ช่องตั้งแต่คอลัมน์ J
'
' ถ้าวันหน้ามีการเพิ่ม/ลบคอลัมน์ในชีต MA&Optimize NER
' ต้องมาแก้ลิสต์นี้ให้ตรงด้วย มิฉะนั้น VerifyExportHeaders จะเตือนและหยุด export
'-------------------------------------------------------------------------------------
Private Function HeaderNames() As Variant
    HeaderNames = Array( _
        "Period", _
        "RO", _
        "Zone", _
        "Province", _
        "ItemCode", _
        "ArtNo", _
        "Description", _
        "Unit", _
        "OnhandBefore", _
        "Alloc", _
        "FromHub", _
        "ExtraAlloc", _
        "Total3", _
        "Reason", _
        "ControlKey", _
        "PeriodSeq", _
        "PrevPeriod", _
        "PrevExist", _
        "PrevBefore", _
        "PrevReceived", _
        "ExpectedBefore", _
        "NetUsed", _
        "DaysPeriod", _
        "BurnRateDay", _
        "DaysCover", _
        "SuggestNextAlloc", _
        "Status", _
        "RiskFlag", _
        "DataSet")
End Function


Private Function BuildHeaders(ByVal lo As ListObject, ByVal nOutCol As Long) As String()
    Dim h()  As String
    Dim def  As Variant
    Dim c    As Long
    Dim n    As Long

    def = HeaderNames()

    ReDim h(1 To nOutCol)
    n = lo.ListColumns.Count

    For c = 1 To n
        If c <= UBound(def) + 1 Then
            h(c) = CStr(def(c - 1))
        Else
            h(c) = SafeHeader(lo.ListColumns(c).Name, c)
        End If
    Next c

    h(n + 1) = "CutoffDate"
    h(n + 2) = "ExportedAt"
    BuildHeaders = h
End Function


Private Function SafeHeader(ByVal s As String, ByVal idx As Long) As String
    Dim i  As Long
    Dim ch As String
    Dim r  As String

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") _
           Or (ch >= "0" And ch <= "9") Or ch = "_" Then
            r = r & ch
        End If
    Next i

    If Len(r) = 0 Then r = "Col" & idx
    SafeHeader = r
End Function


Private Function CleanCell(ByVal v As Variant) As Variant
    If IsError(v) Then
        CleanCell = ""
    ElseIf IsNull(v) Then
        CleanCell = ""
    ElseIf IsEmpty(v) Then
        CleanCell = ""
    Else
        CleanCell = v
    End If
End Function


Private Function IsoDate(ByVal d As Object, ByVal per As String) As String
    Dim k As String
    If d Is Nothing Then Exit Function
    k = LCase$(Trim$(per))
    If Len(k) = 0 Then Exit Function
    If Not d.Exists(k) Then Exit Function
    On Error Resume Next
    IsoDate = Format$(CDate(d(k)), "yyyy-mm-dd")
    On Error GoTo 0
End Function


Private Function CutoffDict() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    AddPeriods d, modRefreshMA.TBL_PER_MA
    AddPeriods d, modRefreshMA.TBL_PER_OPT
    Set CutoffDict = d
End Function


Private Sub AddPeriods(ByVal d As Object, ByVal tblName As String)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim r  As Long
    Dim k  As String
    Dim v  As Variant

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(modRefreshMA.SH_PER)
    If Not ws Is Nothing Then Set lo = ws.ListObjects(tblName)
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub
    If lo.ListRows.Count = 0 Then Exit Sub

    For r = 1 To lo.ListRows.Count
        k = LCase$(Trim$(modRefreshMA.SafeTxt(lo.ListRows(r).Range.Cells(1, 1).Value)))
        v = lo.ListRows(r).Range.Cells(1, 2).Value
        If Len(k) > 0 And Not d.Exists(k) Then
            If IsDate(v) Or IsNumeric(v) Then d.Add k, v
        End If
    Next r
End Sub


Private Function EnsureExportSheet() As Worksheet
    Dim ws  As Worksheet
    Dim evt As Boolean

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_EXPORT)
    On Error GoTo 0

    If ws Is Nothing Then
        evt = Application.EnableEvents
        Application.EnableEvents = False
        Set ws = ThisWorkbook.Worksheets.Add( _
                    After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SH_EXPORT
        Application.EnableEvents = evt
    End If

    Set EnsureExportSheet = ws
End Function


Private Sub ResizeExportTable(ByVal ws As Worksheet, ByVal nRowsTotal As Long, _
                              ByVal nCols As Long)
    Dim lo As ListObject

    If nRowsTotal < 2 Then nRowsTotal = 2

    On Error Resume Next
    Set lo = ws.ListObjects(TBL_EXPORT)
    On Error GoTo 0

    If lo Is Nothing Then
        Set lo = ws.ListObjects.Add(xlSrcRange, _
                    ws.Range("A1").Resize(nRowsTotal, nCols), , xlYes)
        lo.Name = TBL_EXPORT
        lo.TableStyle = "TableStyleMedium2"
    Else
        lo.Resize ws.Range("A1").Resize(nRowsTotal, nCols)
    End If
End Sub


Private Sub ClearBelow(ByVal ws As Worksheet, ByVal fromRow As Long, ByVal nCols As Long)
    Dim lastR As Long
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR > fromRow Then
        ws.Range(ws.Cells(fromRow + 1, 1), ws.Cells(lastR, nCols)).Clear
    End If
End Sub
