' Automated integration test. No prompts, external programs, or APIs in VBA.
Public Function CVCarrierSelfTest(ByVal testDirectory As String) As String
    Dim ws As Worksheet, wb As Workbook, reloaded As Workbook, r As Range, c As Range
    Dim cs As Collection, data() As Byte, enc() As Byte, packetBytes() As Byte, decoded() As Byte
    Dim priv() As Byte, pub() As Byte, privateText As String, publicText As String
    Dim i As Long, row As Long, col As Long, original As Variant, cellCount As Long, fail As String
    Dim phase As String, path As String, n As Long, j As Long, bits As Long, testMode As Long
    On Error GoTo Failed
    phase = "Create source"
    Set ws = ThisWorkbook.Worksheets(1): ws.Name = SCR_SHEET
    ws.Range("A5").Value2 = ChrW(1055) & ChrW(1088) & ChrW(1080) & ChrW(1074) & ChrW(1077) & ChrW(1090) & " " & ChrW(-10179) & ChrW(-8704)
    ws.Range("B5").NumberFormat = "@": ws.Range("B5").Value2 = "=literal"
    ws.Range("C5").Value2 = 123.456
    ws.Range("A6").Value2 = True
    ws.Range("B6").Value2 = False
    ws.Range("C6").Value2 = "a" & vbCrLf & "b:c"
    ws.Range("A7").Value2 = 45292: ws.Range("A7").NumberFormat = "yyyy-mm-dd"
    ws.Range("B7").Formula = "=1+2"
    ws.Range("C9").Value2 = "last"
    Set r = ws.Range("A5:C9")
    data = Serialize(r)
    phase = "Asymmetric encryption"
    testMode = ModeNumber()
    If testMode = 1 Then
        priv = CVDemoPrivate()
    Else
        priv = CVRandom(32)
    End If
    pub = CVPublic(priv, testMode)
    privateText = CVEncode(priv): publicText = CVEncode(pub)
    enc = CVSeal(data, publicText, testMode): packetBytes = Packet(enc, testMode)
    Set wb = Workbooks.Add(xlWBATWorksheet)
    wb.Worksheets(1).Name = HOST_SHEET
    For row = 1 To 100
        For col = 1 To 3
            Set c = wb.Worksheets(1).Cells(row, col)
            c.Value2 = "Ordinary carrier text: report and notes " & CStr(row)
            c.Font.Color = RGB(62, 83, 105): c.Orientation = -7
        Next
    Next
    With wb.Worksheets(1)
        .Range("D1").Value2 = 123
        .Range("D2").Formula = "=1+2"
        .Range("D3:E3").Merge: .Range("D3").Value2 = "Merged heading"
        .Range("D4").Value2 = "Conditional formatting"
        .Range("D4").FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
    End With
    Set cs = CarrierCells(wb.Worksheets(1))
    If cs.Count <> 300 Then Err.Raise 5, , "Carrier filtering failed: " & cs.Count
    original = wb.Worksheets(1).Range("A1:C100").Value2
    phase = "Embed formatting": WriteCarrier cs, packetBytes
    For Each c In wb.Worksheets(1).Range("A1:C100")
        If c.Value2 <> original(c.Row, c.Column) Then Err.Raise 5, , "Carrier text changed"
        For j = 1 To Len(c.Value2)
            n = CLng(c.Characters(j, 1).Font.Color)
            If Abs((n And 255) - 62) > 3 Or Abs(((n \ 256) And 255) - 83) > 3 Or Abs(((n \ 65536) And 255) - 105) > 3 Then Err.Raise 5, , "Color difference exceeds 3"
        Next
    Next
    phase = "Save and reopen XLSX"
    path = testDirectory & "\carrier-roundtrip.xlsx"
    wb.SaveAs Filename:=path, FileFormat:=xlOpenXMLWorkbook
    wb.Close SaveChanges:=False: Set wb = Nothing
    Set reloaded = Workbooks.Open(path, UpdateLinks:=0, ReadOnly:=True)
    Set cs = CarrierCells(reloaded.Worksheets(HOST_SHEET))
    enc = ReadCarrier(cs)
    phase = "Decrypt and restore range"
    decoded = CVOpen(enc, privateText, testMode)
    If CVEncode(data) <> CVEncode(decoded) Then Err.Raise 5, , "Byte round trip differs"
    RestoreRange decoded
    For row = 1 To r.Rows.Count
        For col = 1 To r.Columns.Count
            Set c = ws.Range(OUTPUT_CELL).Offset(row - 1, col - 1)
            If VarType(c.Value2) <> VarType(r.Cells(row, col).Value2) Then Err.Raise 5, , "Type changed at " & c.Address
            If c.Value2 <> r.Cells(row, col).Value2 Then Err.Raise 5, , "Value changed at " & c.Address
            If c.NumberFormat <> r.Cells(row, col).NumberFormat Then Err.Raise 5, , "Number format changed"
        Next
    Next
    If ws.Range("E1").HasFormula Then Err.Raise 5, , "Text executed as formula"
    phase = "Negative cases"
    ws.Protect
    On Error Resume Next
    Err.Clear: RestoreRange decoded
    n = Err.Number: Err.Clear
    On Error GoTo Failed
    ws.Unprotect
    If n = 0 Then Err.Raise 5, , "Protected output accepted"
    ws.Range("D1:E1").Merge
    On Error Resume Next
    Err.Clear: RestoreRange decoded
    n = Err.Number: Err.Clear
    On Error GoTo Failed
    ws.Range("D1:E1").UnMerge
    If n = 0 Then Err.Raise 5, , "Merged output accepted"
    packetBytes(0) = 0
    WriteCarrier cs, packetBytes
    On Error Resume Next
    Err.Clear: enc = ReadCarrier(cs)
    n = Err.Number: Err.Clear
    On Error GoTo Failed
    If n = 0 Then Err.Raise 5, , "Invalid magic accepted"
    packetBytes(0) = 67
    PutLength packetBytes, 6, MAX_BYTES + 1
    WriteCarrier cs, packetBytes
    On Error Resume Next
    Err.Clear: enc = ReadCarrier(cs)
    n = Err.Number: Err.Clear
    On Error GoTo Failed
    If n = 0 Then Err.Raise 5, , "Invalid length accepted"
    PutLength packetBytes, 6, UBound(packetBytes) - 9
    packetBytes(4) = 3 - testMode
    WriteCarrier cs, packetBytes
    On Error Resume Next
    Err.Clear: enc = ReadCarrier(cs)
    n = Err.Number: Err.Clear
    On Error GoTo Failed
    If n = 0 Then Err.Raise 5, , "Mode mismatch accepted"
    reloaded.Close SaveChanges:=False: Set reloaded = Nothing
    CVCarrierSelfTest = "PASS: carrier XLSX round trip, Unicode, typed range, literal formulas, date formats, colors, skipped cells, mode rejection"
    Exit Function
Failed:
    fail = phase & ": " & Err.Number & " " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    If Not reloaded Is Nothing Then reloaded.Close SaveChanges:=False
    CVCarrierSelfTest = "FAIL: " & fail
End Function
