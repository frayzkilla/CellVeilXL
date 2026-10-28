Attribute VB_Name = "cellveilxl"
Option Explicit
' CellVeilXL 2.0: self-contained module; import only this file into Excel.
Public Const CRYPTO_MODE As String = "production"  ' "demo" or "production"
Private Const SCR_SHEET As String = "List1"
Private Const HOST_SHEET As String = "List2"
Private Const SOURCE_RANGE As String = "A1:A10"
Private Const PUBLIC_KEY_CELL As String = "B1"
Private Const PRIVATE_KEY_CELL As String = "B2"
Private Const OUTPUT_CELL As String = "D1"
Private Const HOST_RANGE As String = "" ' Empty = UsedRange; same selection for extraction
Private Const MAX_BYTES As Long = 1048576
Private Const MAX_CELLS As Long = 200000
Private Const FONTS As String = "Arial|Calibri|Segoe UI|Tahoma"
Private gStep As String
Private rngStarted As Boolean
Private sb(255) As Byte
Private sbReady As Boolean

Private Function ModeNumber() As Long
    Select Case LCase$(CRYPTO_MODE)
        Case "demo": ModeNumber = 1
        Case "production": ModeNumber = 2
        Case Else: Err.Raise 5, , "CRYPTO_MODE must be demo or production"
    End Select
End Function
Private Function CheckedRange(ByVal ws As Worksheet, ByVal address As String) As Range
    Dim r As Range
    Set r = ws.Range(address)
    If r.Areas.Count <> 1 Or r.Cells.CountLarge > MAX_CELLS Then Err.Raise 5, , "Use one rectangle with at most 200000 cells"
    Set CheckedRange = r
End Function
Private Function Source() As Range
    Set Source = CheckedRange(ThisWorkbook.Worksheets(SCR_SHEET), SOURCE_RANGE)
End Function
Private Function Host(ByVal ws As Worksheet) As Range
    Dim r As Range
    If Len(HOST_RANGE) = 0 Then
        Set r = ws.UsedRange
    Else
        Set r = ws.Range(HOST_RANGE)
    End If
    If r.Areas.Count <> 1 Or r.Cells.CountLarge > MAX_CELLS Then Err.Raise 5, , "Carrier must be one rectangle with at most 200000 cells"
    Set Host = r
End Function
Private Function Usable(ByVal c As Range) As Boolean
    If c.MergeCells Or c.HasFormula Or IsError(c.Value2) Then Exit Function
    If VarType(c.Value2) <> vbString Then Exit Function
    If Len(c.Value2) = 0 Then Exit Function
    If c.Font.Bold Then Exit Function
    If c.FormatConditions.Count > 0 Then Exit Function
    If c.Orientation <> xlHorizontal Then
        If c.Orientation < -90 Or c.Orientation > 90 Then Exit Function
    End If
    Usable = True
End Function
Private Function CarrierCells(ByVal ws As Worksheet) As Collection
    Dim result As New Collection, c As Range
    For Each c In Host(ws).Cells
        If Usable(c) Then result.Add c
    Next
    Set CarrierCells = result
End Function
Private Function Capacity(ByVal cs As Collection) As Double
    Dim c As Range, n As Double
    For Each c In cs: n = n + 4# + 6# * Len(c.Value2): Next
    Capacity = n
End Function
Private Function Field(ByVal s As String) As String
    Field = CStr(Len(s)) & ":" & s
End Function
Private Function ReadField(ByVal s As String, ByRef pos As Long) As String
    Dim colon As Long, n As Double, token As String, i As Long
    colon = InStr(pos, s, ":")
    If colon = 0 Or colon - pos > 8 Then Err.Raise 5, , "Invalid range record"
    token = Mid$(s, pos, colon - pos)
    If Len(token) = 0 Then Err.Raise 5, , "Empty field length"
    For i = 1 To Len(token)
        If Mid$(token, i, 1) < "0" Or Mid$(token, i, 1) > "9" Then Err.Raise 5, , "Invalid field length"
    Next
    n = Val(token)
    If n > Len(s) - colon Then Err.Raise 5, , "Truncated range record"
    ReadField = Mid$(s, colon + 1, CLng(n)): pos = colon + CLng(n) + 1
End Function
Private Function ToBytes(ByVal s As String) As Byte()
    Dim b() As Byte, i As Long, n As Long
    If Len(s) * 2# > MAX_BYTES - 64 Then Err.Raise 5, , "Source exceeds one MiB packet limit"
    ReDim b(0 To Len(s) * 2 - 1)
    For i = 1 To Len(s)
        n = AscW(Mid$(s, i, 1)) And 65535
        b((i - 1) * 2) = n And 255: b((i - 1) * 2 + 1) = n \ 256
    Next
    ToBytes = b
End Function
Private Function FromBytes(ByRef b() As Byte) As String
    Dim parts() As String, i As Long, n As Long, size As Long
    size = UBound(b) + 1
    If size Mod 2 <> 0 Then Err.Raise 5, , "Invalid UTF-16 payload"
    ReDim parts(0 To size \ 2 - 1)
    For i = 0 To size - 1 Step 2
        n = CLng(b(i)) + 256& * b(i + 1)
        If n > 32767 Then n = n - 65536
        parts(i \ 2) = ChrW(n)
    Next
    FromBytes = Join(parts, "")
End Function
Private Function Serialize(ByVal r As Range) As Byte()
    Dim records() As String, c As Range, i As Long, kind As String, value As String, s As String
    ReDim records(0 To CLng(r.Cells.CountLarge) - 1)
    For Each c In r.Cells
        If c.MergeCells Then Err.Raise 5, , "Source includes merged cells"
        If IsError(c.Value2) Then Err.Raise 5, , "Source includes an Excel error: " & c.Address
        Select Case VarType(c.Value2)
            Case vbEmpty: kind = "z": value = ""
            Case vbString: kind = "s": value = CStr(c.Value2)
            Case vbBoolean
                kind = "b": value = "0": If c.Value2 Then value = "1"
            Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal
                kind = "n": value = Trim$(Str$(c.Value2))
            Case Else: Err.Raise 5, , "Unsupported source cell: " & c.Address
        End Select
        records(i) = Field(kind) & Field(value) & Field(CStr(c.NumberFormat)): i = i + 1
    Next
    s = Field(CStr(r.Rows.Count)) & Field(CStr(r.Columns.Count)) & Join(records, "")
    Serialize = ToBytes(s)
End Function
Private Sub RestoreRange(ByRef data() As Byte)
    Dim s As String, pos As Long, h As Long, w As Long, count As Long, i As Long
    Dim kinds() As String, values() As Variant, formats() As String, v As String, token As String
    Dim anchor As Range, target As Range, c As Range, oldValues() As Variant, oldFormulas() As Variant, oldHasFormula() As Boolean, oldFormats() As String
    Dim num As Long, msg As String, started As Boolean
    s = FromBytes(data): pos = 1
    token = ReadField(s, pos): h = ValidDimension(token, 1048576)
    token = ReadField(s, pos): w = ValidDimension(token, 16384)
    If CDbl(h) * w > MAX_CELLS Then Err.Raise 5, , "Decoded range is too large"
    count = h * w
    ReDim kinds(0 To count - 1): ReDim values(0 To count - 1): ReDim formats(0 To count - 1)
    For i = 0 To count - 1
        kinds(i) = ReadField(s, pos): v = ReadField(s, pos): formats(i) = ReadField(s, pos)
        If Len(v) > 32767 Or Len(formats(i)) > 255 Then Err.Raise 5, , "Invalid decoded cell length"
        Select Case kinds(i)
            Case "z": If Len(v) <> 0 Then Err.Raise 5, , "Invalid empty cell"
            Case "s": values(i) = v
            Case "b"
                If v <> "0" And v <> "1" Then Err.Raise 5, , "Invalid Boolean"
                values(i) = (v = "1")
            Case "n"
                If Not ValidNumber(v) Then Err.Raise 5, , "Invalid number"
                values(i) = Val(v)
            Case Else: Err.Raise 5, , "Invalid cell type (wrong key or damaged data)"
        End Select
    Next
    If pos <> Len(s) + 1 Then Err.Raise 5, , "Trailing range data"
    Set anchor = ThisWorkbook.Worksheets(SCR_SHEET).Range(OUTPUT_CELL)
    If anchor.Cells.CountLarge <> 1 Then Err.Raise 5, , "OUTPUT_CELL must be one cell"
    If anchor.Row + h - 1 > 1048576 Or anchor.Column + w - 1 > 16384 Then Err.Raise 5, , "Output exceeds worksheet bounds"
    Set target = anchor.Resize(h, w)
    If target.Worksheet.ProtectContents Then Err.Raise 5, , "Output sheet is protected"
    If Not Intersect(target, target.Worksheet.Range(PUBLIC_KEY_CELL & "," & PRIVATE_KEY_CELL)) Is Nothing Then Err.Raise 5, , "Output overlaps key cells"
    ReDim oldFormats(0 To count - 1)
    ReDim oldValues(0 To count - 1): ReDim oldFormulas(0 To count - 1): ReDim oldHasFormula(0 To count - 1)
    i = 0
    For Each c In target.Cells
        If c.MergeCells Or c.HasArray Then Err.Raise 5, , "Output contains merged cells or array formulas"
        oldFormats(i) = CStr(c.NumberFormat)
        oldValues(i) = c.Value2: oldFormulas(i) = c.Formula: oldHasFormula(i) = c.HasFormula
        i = i + 1
    Next
    On Error GoTo Rollback
    started = True: i = 0
    For Each c In target.Cells
        ' Text must never be interpreted as an Excel formula.
        If kinds(i) = "s" Then
            c.NumberFormat = "@": c.Value2 = values(i)
        ElseIf kinds(i) = "z" Then
            c.ClearContents
        Else
            c.Value2 = values(i)
        End If
        c.NumberFormat = formats(i): i = i + 1
    Next
    Exit Sub
Rollback:
    num = Err.Number: msg = Err.Description
    On Error Resume Next
    i = 0
    For Each c In target.Cells
        If oldHasFormula(i) Then
            c.Formula = oldFormulas(i)
        Else
            c.NumberFormat = "@": c.Value2 = oldValues(i)
        End If
        c.NumberFormat = oldFormats(i): i = i + 1
    Next
    On Error GoTo 0
    Err.Raise num, , msg
End Sub
Private Function ValidDimension(ByVal s As String, ByVal maximum As Long) As Long
    Dim i As Long, n As Double
    If Len(s) = 0 Or Len(s) > 7 Then Err.Raise 5, , "Invalid range dimensions"
    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Err.Raise 5, , "Invalid range dimensions"
    Next
    n = Val(s)
    If n < 1 Or n > maximum Then Err.Raise 5, , "Invalid range dimensions"
    ValidDimension = CLng(n)
End Function
Private Function ValidNumber(ByVal s As String) As Boolean
    Dim i As Long, ch As String
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If InStr(1, "0123456789.Ee+-", ch, vbBinaryCompare) = 0 Then Exit Function
    Next
    ' Str(Val()) normalization rejects partially parsed or malformed strings.
    On Error GoTo Bad
    ValidNumber = (Val(s) = CDbl(Replace(s, ".", Application.International(xlDecimalSeparator))))
Bad:
End Function
Private Sub PutLength(ByRef b() As Byte, ByVal offset As Long, ByVal n As Long)
    Dim i As Long
    For i = 0 To 3: b(offset + i) = n And 255: n = n \ 256: Next
End Sub
Private Function Packet(ByRef encrypted() As Byte, ByVal mode As Long) As Byte()
    Dim b() As Byte, n As Long, i As Long
    n = UBound(encrypted) + 1
    If n > MAX_BYTES Then Err.Raise 5, , "Encrypted packet too large"
    ReDim b(0 To n + 9)
    b(0) = 67: b(1) = 86: b(2) = 88: b(3) = 51: b(4) = mode: b(5) = 1
    PutLength b, 6, n
    For i = 0 To n - 1: b(10 + i) = encrypted(i): Next
    Packet = b
End Function
Private Function TakeBits(ByRef b() As Byte, ByRef pos As Long, ByVal width As Long) As Long
    Dim i As Long, v As Long, total As Long
    total = (UBound(b) + 1) * 8
    For i = 1 To width
        v = v * 2
        If pos < total Then v = v Or ((b(pos \ 8) \ (2 ^ (7 - (pos And 7)))) And 1)
        pos = pos + 1
    Next
    TakeBits = v
End Function
Private Sub WriteCarrier(ByVal cs As Collection, ByRef b() As Byte)
    Dim c As Range, pos As Long, total As Long, i As Long, n As Long, color As Long
    Dim red As Long, green As Long, blue As Long, names As Variant, baseValue As Long
    total = (UBound(b) + 1) * 8: names = Split(FONTS, "|")
    If Capacity(cs) < total Then Err.Raise 5, , "Insufficient carrier: need " & total & " bits; available " & Capacity(cs)
    For Each c In cs
        If pos >= total Then Exit For
        baseValue = 0
        If c.Orientation <> xlHorizontal Then baseValue = Int(CDbl(c.Orientation) / 4#) * 4
        If baseValue < -88 Then baseValue = -88
        If baseValue > 84 Then baseValue = 84
        c.Orientation = baseValue + TakeBits(b, pos, 2)
        c.Font.Name = names(TakeBits(b, pos, 2))
        For i = 1 To Len(c.Value2)
            If pos >= total Then Exit For
            n = TakeBits(b, pos, 6): color = CLng(c.Characters(i, 1).Font.Color)
            red = (color And 255): green = (color \ 256) And 255: blue = (color \ 65536) And 255
            red = (red And 252) Or (n \ 16)
            green = (green And 252) Or ((n \ 4) And 3)
            blue = (blue And 252) Or (n And 3)
            c.Characters(i, 1).Font.Color = RGB(red, green, blue)
        Next
    Next
End Sub
Private Function Feed(ByRef b() As Byte, ByRef count As Long, ByRef acc As Long, ByRef bits As Long, ByVal value As Long, ByVal width As Long, ByRef needed As Long, ByVal available As Double) As Boolean
    Dim n As Double
    acc = acc * (2 ^ width) + value: bits = bits + width
    Do While bits >= 8
        bits = bits - 8: b(count) = acc \ (2 ^ bits): count = count + 1
        acc = acc And ((2 ^ bits) - 1)
        If count = 10 Then
            If b(0) <> 67 Or b(1) <> 86 Or b(2) <> 88 Or b(3) <> 51 Or b(5) <> 1 Then Err.Raise 5, , "No CellVeilXL v3 packet"
            If b(4) <> ModeNumber() Then Err.Raise 5, , "Packet mode differs from CRYPTO_MODE"
            n = CDbl(b(6)) + 256# * b(7) + 65536# * b(8) + 16777216# * b(9)
            If n < 28 Or n > MAX_BYTES Or (n + 10#) * 8# > available Then Err.Raise 5, , "Invalid packet length"
            needed = CLng(n) + 10: ReDim Preserve b(0 To needed - 1)
        End If
        If needed > 0 Then
            If count = needed Then Feed = True: Exit Function
        End If
    Loop
End Function
Private Function ReadCarrier(ByVal cs As Collection) As Byte()
    Dim b() As Byte, count As Long, acc As Long, bits As Long, needed As Long, available As Double
    Dim c As Range, i As Long, j As Long, n As Long, color As Long, names As Variant
    available = Capacity(cs)
    If available < 80 Then Err.Raise 5, , "Carrier too small for header"
    ReDim b(0 To 9): names = Split(FONTS, "|")
    For Each c In cs
        n = CLng(c.Orientation) And 3
        If Feed(b, count, acc, bits, n, 2, needed, available) Then GoTo Done
        n = -1
        For j = 0 To 3
            If StrComp(CStr(c.Font.Name), CStr(names(j)), vbTextCompare) = 0 Then n = j: Exit For
        Next
        If n < 0 Then Err.Raise 5, , "Unknown carrier font at " & c.Address
        If Feed(b, count, acc, bits, n, 2, needed, available) Then GoTo Done
        For i = 1 To Len(c.Value2)
            color = CLng(c.Characters(i, 1).Font.Color)
            n = (color And 3) * 16 + ((color \ 256) And 3) * 4 + ((color \ 65536) And 3)
            If Feed(b, count, acc, bits, n, 6, needed, available) Then GoTo Done
        Next
    Next
    Err.Raise 5, , "Truncated carrier"
Done:
    ReadCarrier = CVSlice(b, 10, needed - 10)
End Function
Public Sub EmbedScr()
    Dim wsS As Worksheet, wsH As Worksheet, outputBook As Workbook, destination As Worksheet
    Dim data() As Byte, encrypted() As Byte, b() As Byte, cs As Collection, r As Range
    Dim mode As Long, oldScreen As Boolean, oldEvents As Boolean, stateSaved As Boolean, msg As String
    On Error GoTo Failed
    gStep = "Validate configuration": mode = ModeNumber()
    Set wsS = ThisWorkbook.Worksheets(SCR_SHEET): Set wsH = ThisWorkbook.Worksheets(HOST_SHEET)
    If wsS Is wsH Then Err.Raise 5, , "Source and carrier sheets must be different"
    Set r = Source()
    If Not Intersect(r, wsS.Range(PUBLIC_KEY_CELL & "," & PRIVATE_KEY_CELL)) Is Nothing Then Err.Raise 5, , "Source range overlaps key cells"
    gStep = "Serialize and encrypt range"
    data = Serialize(r): encrypted = CVSeal(data, CStr(wsS.Range(PUBLIC_KEY_CELL).Value2), mode)
    b = Packet(encrypted, mode)
    Set cs = CarrierCells(wsH)
    If Capacity(cs) < (UBound(b) + 1) * 8# Then Err.Raise 5, , "Insufficient carrier: need " & (UBound(b) + 1) * 8# & " bits; available " & Capacity(cs)
    oldScreen = Application.ScreenUpdating: oldEvents = Application.EnableEvents: stateSaved = True
    Application.ScreenUpdating = False: Application.EnableEvents = False
    gStep = "Create document copy"
    ' A fresh workbook excludes source sheet, VBA, names, comments and linked objects.
    Set outputBook = Workbooks.Add(xlWBATWorksheet): Set destination = outputBook.Worksheets(1)
    destination.Name = HOST_SHEET
    ' Keep values, formatting and layout of the visible carrier document.
    wsH.UsedRange.Copy
    destination.Range(wsH.UsedRange.Address).PasteSpecial xlPasteValues
    destination.Range(wsH.UsedRange.Address).PasteSpecial xlPasteFormats
    destination.Range(wsH.UsedRange.Address).PasteSpecial xlPasteColumnWidths
    Application.CutCopyMode = False
    Dim row As Range, originalCell As Range
    For Each originalCell In wsH.UsedRange.Cells
        If Not originalCell.MergeCells And VarType(originalCell.Value2) = vbString Then
            With destination.Range(originalCell.Address)
                .NumberFormat = "@": .Value2 = originalCell.Value2
                .NumberFormat = originalCell.NumberFormat
            End With
        End If
    Next
    For Each row In wsH.UsedRange.Rows
        destination.Rows(row.Row).RowHeight = row.RowHeight
        destination.Rows(row.Row).Hidden = row.EntireRow.Hidden
    Next
    For Each row In wsH.UsedRange.Columns
        destination.Columns(row.Column).Hidden = row.EntireColumn.Hidden
    Next
    ' Formula results can become text: keep them out of the carrier in both directions.
    ' Encoding uses the final workbook's cell set, and rechecks its capacity.
    Set cs = CarrierCells(destination)
    gStep = "Write formatting bits": WriteCarrier cs, b
    Application.ScreenUpdating = oldScreen: Application.EnableEvents = oldEvents
    outputBook.Activate
    MsgBox "Created a carrier-only workbook. Save it as .xlsx. " & vbCrLf & _
        "Mode: " & CRYPTO_MODE & ". Rnd is not cryptographically secure in either mode." & vbCrLf & _
        "Keep the document unchanged until extraction.", vbInformation, "CellVeilXL"
    Exit Sub
Failed:
    msg = Err.Description
    On Error Resume Next
    If Not outputBook Is Nothing Then outputBook.Close SaveChanges:=False
    If stateSaved Then
        Application.ScreenUpdating = oldScreen: Application.EnableEvents = oldEvents
    End If
    Application.CutCopyMode = False
    MsgBox gStep & ": " & msg, vbCritical, "CellVeilXL"
End Sub
Public Sub ExtractScr()
    Dim cs As Collection, encrypted() As Byte, data() As Byte, mode As Long
    Dim selected As Variant, carrier As Workbook, opened As Boolean
    Dim oldEvents As Boolean, stateSaved As Boolean, msg As String
    On Error GoTo Failed
    gStep = "Select carrier workbook": mode = ModeNumber()
    selected = Application.GetOpenFilename("Excel document (*.xlsx),*.xlsx")
    If VarType(selected) = vbBoolean Then Exit Sub
    oldEvents = Application.EnableEvents: stateSaved = True: Application.EnableEvents = False
    ' xlsx excludes VBA; no external links are refreshed.
    Set carrier = Workbooks.Open(CStr(selected), UpdateLinks:=0, ReadOnly:=True): opened = True
    Set cs = CarrierCells(carrier.Worksheets(HOST_SHEET))
    gStep = "Read hidden packet": encrypted = ReadCarrier(cs)
    carrier.Close SaveChanges:=False: opened = False
    gStep = "Decrypt range": data = CVOpen(encrypted, CStr(ThisWorkbook.Worksheets(SCR_SHEET).Range(PRIVATE_KEY_CELL).Value2), mode)
    gStep = "Restore range": RestoreRange data
    Application.EnableEvents = oldEvents
    MsgBox "Range restored at " & SCR_SHEET & "!" & OUTPUT_CELL & ".", vbInformation, "CellVeilXL"
    Exit Sub
Failed:
    msg = Err.Description
    On Error Resume Next
    If opened Then carrier.Close SaveChanges:=False
    If stateSaved Then Application.EnableEvents = oldEvents
    MsgBox gStep & ": " & msg, vbCritical, "CellVeilXL"
End Sub


' --- Cryptography ---
' Pure VBA only. Rnd is NOT a cryptographic RNG. This affects BOTH modes.
Private Function Unsigned(ByVal n As Long) As Double
    Unsigned = n
    If n < 0 Then Unsigned = Unsigned + 4294967296#
End Function
Private Function Signed(ByVal n As Double) As Long
    n = n - Int(n / 4294967296#) * 4294967296#
    If n >= 2147483648# Then n = n - 4294967296#
    Signed = CLng(n)
End Function
Private Function Ror(ByVal x As Long, ByVal n As Long) As Long
    Dim u As Double, low As Double
    u = Unsigned(x): low = u - Int(u / (2# ^ n)) * (2# ^ n)
    Ror = Signed(Int(u / (2# ^ n)) + low * (2# ^ (32 - n)))
End Function
Private Function Shr(ByVal x As Long, ByVal n As Long) As Long
    Shr = CLng(Int(Unsigned(x) / (2# ^ n)))
End Function
Public Function CVRandom(ByVal size As Long) As Byte()
    Dim b() As Byte, i As Long
    If Not rngStarted Then
        Randomize: rngStarted = True
    End If
    ReDim b(0 To size - 1)
    For i = 0 To size - 1: b(i) = Int(Rnd() * 256#): Next
    CVRandom = b
End Function
Public Function CVHash(ByRef data() As Byte) As Byte()
    Dim k As Variant, initial As Variant, h(7) As Long, w(63) As Long
    Dim padded() As Byte, out(31) As Byte, size As Long, total As Long
    Dim i As Long, j As Long, offset As Long, bits As Double
    Dim a As Long, b As Long, c As Long, d As Long, e As Long, f As Long, g As Long, hh As Long
    Dim s0 As Long, s1 As Long, ch As Long, maj As Long, t1 As Long, t2 As Long
    k = Array(&H428A2F98, &H71374491, &HB5C0FBCF, &HE9B5DBA5, &H3956C25B, &H59F111F1, &H923F82A4, &HAB1C5ED5, _
        &HD807AA98, &H12835B01, &H243185BE, &H550C7DC3, &H72BE5D74, &H80DEB1FE, &H9BDC06A7, &HC19BF174, _
        &HE49B69C1, &HEFBE4786, &H0FC19DC6, &H240CA1CC, &H2DE92C6F, &H4A7484AA, &H5CB0A9DC, &H76F988DA, _
        &H983E5152, &HA831C66D, &HB00327C8, &HBF597FC7, &HC6E00BF3, &HD5A79147, &H06CA6351, &H14292967, _
        &H27B70A85, &H2E1B2138, &H4D2C6DFC, &H53380D13, &H650A7354, &H766A0ABB, &H81C2C92E, &H92722C85, _
        &HA2BFE8A1, &HA81A664B, &HC24B8B70, &HC76C51A3, &HD192E819, &HD6990624, &HF40E3585, &H106AA070, _
        &H19A4C116, &H1E376C08, &H2748774C, &H34B0BCB5, &H391C0CB3, &H4ED8AA4A, &H5B9CCA4F, &H682E6FF3, _
        &H748F82EE, &H78A5636F, &H84C87814, &H8CC70208, &H90BEFFFA, &HA4506CEB, &HBEF9A3F7, &HC67178F2)
    initial = Array(&H6A09E667, &HBB67AE85, &H3C6EF372, &HA54FF53A, &H510E527F, &H9B05688C, &H1F83D9AB, &H5BE0CD19)
    For i = 0 To 7: h(i) = initial(i): Next
    size = UBound(data) + 1: total = ((size + 9 + 63) \ 64) * 64
    ReDim padded(0 To total - 1)
    For i = 0 To size - 1: padded(i) = data(i): Next
    padded(size) = 128: bits = CDbl(size) * 8#
    For i = total - 1 To total - 8 Step -1
        padded(i) = bits - Int(bits / 256#) * 256#: bits = Int(bits / 256#)
    Next
    For offset = 0 To total - 1 Step 64
        For i = 0 To 15
            j = offset + i * 4
            w(i) = Signed(padded(j) * 16777216# + padded(j + 1) * 65536# + padded(j + 2) * 256# + padded(j + 3))
        Next
        For i = 16 To 63
            s0 = Ror(w(i - 15), 7) Xor Ror(w(i - 15), 18) Xor Shr(w(i - 15), 3)
            s1 = Ror(w(i - 2), 17) Xor Ror(w(i - 2), 19) Xor Shr(w(i - 2), 10)
            w(i) = Signed(Unsigned(w(i - 16)) + Unsigned(s0) + Unsigned(w(i - 7)) + Unsigned(s1))
        Next
        a = h(0): b = h(1): c = h(2): d = h(3): e = h(4): f = h(5): g = h(6): hh = h(7)
        For i = 0 To 63
            s1 = Ror(e, 6) Xor Ror(e, 11) Xor Ror(e, 25)
            ch = (e And f) Xor ((Not e) And g)
            t1 = Signed(Unsigned(hh) + Unsigned(s1) + Unsigned(ch) + Unsigned(CLng(k(i))) + Unsigned(w(i)))
            s0 = Ror(a, 2) Xor Ror(a, 13) Xor Ror(a, 22)
            maj = (a And b) Xor (a And c) Xor (b And c)
            t2 = Signed(Unsigned(s0) + Unsigned(maj))
            hh = g: g = f: f = e: e = Signed(Unsigned(d) + Unsigned(t1))
            d = c: c = b: b = a: a = Signed(Unsigned(t1) + Unsigned(t2))
        Next
        h(0) = Signed(Unsigned(h(0)) + Unsigned(a)): h(1) = Signed(Unsigned(h(1)) + Unsigned(b))
        h(2) = Signed(Unsigned(h(2)) + Unsigned(c)): h(3) = Signed(Unsigned(h(3)) + Unsigned(d))
        h(4) = Signed(Unsigned(h(4)) + Unsigned(e)): h(5) = Signed(Unsigned(h(5)) + Unsigned(f))
        h(6) = Signed(Unsigned(h(6)) + Unsigned(g)): h(7) = Signed(Unsigned(h(7)) + Unsigned(hh))
    Next
    For i = 0 To 7
        bits = Unsigned(h(i))
        For j = 3 To 0 Step -1
            out(i * 4 + j) = bits - Int(bits / 256#) * 256#: bits = Int(bits / 256#)
        Next
    Next
    CVHash = out
End Function
Private Function GFMul(ByVal a As Long, ByVal b As Long) As Long
    Dim r As Long, i As Long
    For i = 0 To 7
        If (b And 1) <> 0 Then r = r Xor a
        b = b \ 2: a = a * 2
        If a > 255 Then a = a Xor 283
    Next
    GFMul = r
End Function
Private Sub InitSBox()
    Dim i As Long, j As Long, inverse As Long, v As Long, r As Long
    If sbReady Then Exit Sub
    For i = 0 To 255
        inverse = 0
        If i <> 0 Then
            For j = 1 To 255
                If GFMul(i, j) = 1 Then inverse = j: Exit For
            Next
        End If
        v = inverse: r = v
        For j = 1 To 4
            v = ((v * 2) And 255) Or (v \ 128): r = r Xor v
        Next
        sb(i) = r Xor 99
    Next
    sbReady = True
End Sub
Private Function AESKeys(ByRef key() As Byte) As Byte()
    Dim rk(239) As Byte, temp(3) As Byte, n As Long, i As Long, rcon As Long, t As Byte
    InitSBox
    For i = 0 To 31: rk(i) = key(i): Next
    n = 32: rcon = 1
    Do While n < 240
        For i = 0 To 3: temp(i) = rk(n - 4 + i): Next
        If n Mod 32 = 0 Then
            t = temp(0): temp(0) = sb(temp(1)) Xor rcon
            temp(1) = sb(temp(2)): temp(2) = sb(temp(3)): temp(3) = sb(t)
            rcon = GFMul(rcon, 2)
        ElseIf n Mod 32 = 16 Then
            For i = 0 To 3: temp(i) = sb(temp(i)): Next
        End If
        For i = 0 To 3
            rk(n) = rk(n - 32) Xor temp(i): n = n + 1
        Next
    Loop
    AESKeys = rk
End Function
Private Function AESBlock(ByRef block() As Byte, ByRef rk() As Byte) As Byte()
    Dim s(15) As Byte, t(15) As Byte, i As Long, row As Long, col As Long, round As Long
    Dim a As Long, b As Long, c As Long, d As Long
    For i = 0 To 15: s(i) = block(i) Xor rk(i): Next
    For round = 1 To 14
        For col = 0 To 3
            For row = 0 To 3
                t(col * 4 + row) = sb(s(((col + row) Mod 4) * 4 + row))
            Next
        Next
        For i = 0 To 15: s(i) = t(i): Next
        If round < 14 Then
            For col = 0 To 3
                i = col * 4: a = s(i): b = s(i + 1): c = s(i + 2): d = s(i + 3)
                s(i) = GFMul(a, 2) Xor GFMul(b, 3) Xor c Xor d
                s(i + 1) = a Xor GFMul(b, 2) Xor GFMul(c, 3) Xor d
                s(i + 2) = a Xor b Xor GFMul(c, 2) Xor GFMul(d, 3)
                s(i + 3) = GFMul(a, 3) Xor b Xor c Xor GFMul(d, 2)
            Next
        End If
        For i = 0 To 15: s(i) = s(i) Xor rk(round * 16 + i): Next
    Next
    AESBlock = s
End Function
Public Function CVCTR(ByRef data() As Byte, ByRef key() As Byte, ByRef nonce() As Byte) As Byte()
    Dim rk() As Byte, counter() As Byte, stream() As Byte, out() As Byte
    Dim i As Long, j As Long, size As Long
    rk = AESKeys(key): counter = nonce: size = UBound(data) + 1
    ReDim out(0 To size - 1)
    For i = 0 To size - 1 Step 16
        stream = AESBlock(counter, rk)
        For j = 0 To 15
            If i + j < size Then out(i + j) = data(i + j) Xor stream(j)
        Next
        For j = 15 To 0 Step -1
            If counter(j) < 255 Then counter(j) = counter(j) + 1: Exit For
            counter(j) = 0
        Next
    Next
    CVCTR = out
    Erase rk
End Function
Public Function CVJoin(ByRef a() As Byte, ByRef b() As Byte) As Byte()
    Dim r() As Byte, i As Long, n As Long
    n = UBound(a) + 1: ReDim r(0 To n + UBound(b))
    For i = 0 To n - 1: r(i) = a(i): Next
    For i = 0 To UBound(b): r(n + i) = b(i): Next
    CVJoin = r
End Function
Public Function CVSlice(ByRef data() As Byte, ByVal start As Long, ByVal size As Long) As Byte()
    Dim r() As Byte, i As Long
    If size < 1 Or start < 0 Or start + size > UBound(data) + 1 Then Err.Raise 5, , "Invalid byte slice"
    ReDim r(0 To size - 1)
    For i = 0 To size - 1: r(i) = data(start + i): Next
    CVSlice = r
End Function
Public Function CVAscii(ByVal value As String) As Byte()
    Dim b() As Byte, i As Long
    ReDim b(0 To Len(value) - 1)
    For i = 1 To Len(value): b(i - 1) = AscW(Mid$(value, i, 1)): Next
    CVAscii = b
End Function
Private Function HMAC(ByRef key() As Byte, ByRef data() As Byte) As Byte()
    Dim ip(63) As Byte, op(63) As Byte, i As Long, tmp() As Byte
    For i = 0 To 63
        ip(i) = 54: op(i) = 92
        If i <= UBound(key) Then
            ip(i) = ip(i) Xor key(i): op(i) = op(i) Xor key(i)
        End If
    Next
    tmp = CVJoin(ip, data): tmp = CVHash(tmp)
    tmp = CVJoin(op, tmp): HMAC = CVHash(tmp)
End Function
Public Function CVKDF(ByRef secretValue() As Byte, ByRef ep() As Byte, ByRef recipient() As Byte, ByVal mode As Long) As Byte()
    Dim salt(31) As Byte, prk() As Byte, info() As Byte, tmp() As Byte
    prk = HMAC(salt, secretValue)
    info = CVAscii("CellVeilXL-v3/" & CStr(mode))
    tmp = CVJoin(info, ep): tmp = CVJoin(tmp, recipient)
    info = CVAscii(Chr$(1)): tmp = CVJoin(tmp, info)
    CVKDF = HMAC(prk, tmp)
End Function
Public Function CVEncode(ByRef b() As Byte) As String
    Dim alphabet As String, out As String, acc As Long, bits As Long, i As Long
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_"
    For i = 0 To UBound(b)
        acc = acc * 256 + b(i): bits = bits + 8
        Do While bits >= 6
            bits = bits - 6
            out = out & Mid$(alphabet, (acc \ (2 ^ bits)) + 1, 1)
            acc = acc And ((2 ^ bits) - 1)
        Loop
    Next
    If bits > 0 Then out = out & Mid$(alphabet, acc * (2 ^ (6 - bits)) + 1, 1)
    CVEncode = out
End Function
Public Function CVDecode(ByVal value As String, ByVal size As Long) As Byte()
    Dim alphabet As String, b() As Byte, i As Long, n As Long, acc As Long, bits As Long, j As Long
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_"
    value = Replace(Replace(Replace(Trim$(value), " ", ""), vbCr, ""), vbLf, "")
    If Len(value) <> (size * 8 + 5) \ 6 Then Err.Raise 5, , "Key length does not match selected mode"
    ReDim b(0 To size - 1)
    For i = 1 To Len(value)
        n = InStr(1, alphabet, Mid$(value, i, 1), vbBinaryCompare) - 1
        If n < 0 Then Err.Raise 5, , "Invalid Base64URL character"
        acc = acc * 64 + n: bits = bits + 6
        If bits >= 8 Then
            bits = bits - 8
            If j >= size Then Err.Raise 5, , "Invalid key length"
            b(j) = acc \ (2 ^ bits): j = j + 1
            acc = acc And ((2 ^ bits) - 1)
        End If
    Next
    If j <> size Or acc <> 0 Then Err.Raise 5, , "Non-canonical key encoding"
    CVDecode = b
End Function

Public Function CVPublic(ByRef privateKey() As Byte, ByVal mode As Long) As Byte()
    Dim baseValue(31) As Byte
    If mode = 1 Then
        CVPublic = CVDemoPublic(privateKey)
    ElseIf mode = 2 Then
        baseValue(0) = 9: CVPublic = CVX25519(privateKey, baseValue)
    Else
        Err.Raise 5, , "Unknown crypto mode"
    End If
End Function
Private Function SharedKey(ByRef privateKey() As Byte, ByRef publicKey() As Byte, ByVal mode As Long) As Byte()
    Dim b() As Byte, i As Long, nonzero As Long
    If mode = 1 Then
        b = CVDemoShared(privateKey, publicKey)
    Else
        b = CVX25519(privateKey, publicKey)
    End If
    For i = 0 To UBound(b): nonzero = nonzero Or b(i): Next
    If nonzero = 0 Then Err.Raise 5, , "Invalid all-zero secretValue secret"
    SharedKey = b
End Function
Public Function CVSeal(ByRef data() As Byte, ByVal publicText As String, ByVal mode As Long) As Byte()
    Dim priv() As Byte, recipient() As Byte, ep() As Byte, secretValue() As Byte
    Dim key() As Byte, nonce() As Byte, encrypted() As Byte, r() As Byte
    If mode = 1 Then
        recipient = CVDecode(publicText, 11): priv = CVDemoPrivate()
    ElseIf mode = 2 Then
        recipient = CVDecode(publicText, 32): priv = CVRandom(32)
    Else
        Err.Raise 5, , "Unknown crypto mode"
    End If
    ep = CVPublic(priv, mode): secretValue = SharedKey(priv, recipient, mode)
    key = CVKDF(secretValue, ep, recipient, mode): nonce = CVRandom(16)
    encrypted = CVCTR(data, key, nonce)
    r = CVJoin(ep, nonce): CVSeal = CVJoin(r, encrypted)
    Erase priv: Erase secretValue: Erase key
End Function
Public Function CVOpen(ByRef data() As Byte, ByVal privateText As String, ByVal mode As Long) As Byte()
    Dim priv() As Byte, recipient() As Byte, ep() As Byte, secretValue() As Byte
    Dim key() As Byte, nonce() As Byte, encrypted() As Byte, size As Long
    If mode = 1 Then
        size = 11
    ElseIf mode = 2 Then
        size = 32
    Else
        Err.Raise 5, , "Unknown crypto mode"
    End If
    priv = CVDecode(privateText, size): recipient = CVPublic(priv, mode)
    ep = CVSlice(data, 0, size): nonce = CVSlice(data, size, 16)
    secretValue = SharedKey(priv, ep, mode): key = CVKDF(secretValue, ep, recipient, mode)
    encrypted = CVSlice(data, size + 16, UBound(data) + 1 - size - 16)
    CVOpen = CVCTR(encrypted, key, nonce)
    Erase priv: Erase secretValue: Erase key
End Function

' --- Asymmetric field arithmetic ---
' Finite-field X25519 arithmetic, radix 2^16. Based on the public-domain
' TweetNaCl Montgomery ladder. Not a constant-time VBA implementation.
Private Sub Carry(ByRef o As Variant)
    Dim i As Long, c As Double
    For i = 0 To 15
        o(i) = o(i) + 65536#
        c = Int(o(i) / 65536#)
        o(i) = o(i) - c * 65536#
        If i < 15 Then
            o(i + 1) = o(i + 1) + c - 1
        Else
            o(0) = o(0) + 38# * (c - 1)
        End If
    Next
End Sub
Private Function ZeroField() As Variant
    Dim a(15) As Double
    ZeroField = a
End Function
Private Function FAdd(ByVal a As Variant, ByVal b As Variant) As Variant
    Dim o As Variant, i As Long
    o = ZeroField()
    For i = 0 To 15: o(i) = a(i) + b(i): Next
    FAdd = o
End Function
Private Function FSub(ByVal a As Variant, ByVal b As Variant) As Variant
    Dim o As Variant, i As Long
    o = ZeroField()
    For i = 0 To 15: o(i) = a(i) - b(i): Next
    FSub = o
End Function
Private Function FMul(ByVal a As Variant, ByVal b As Variant) As Variant
    Dim t(30) As Double, o As Variant, i As Long, j As Long
    For i = 0 To 15
        For j = 0 To 15: t(i + j) = t(i + j) + a(i) * b(j): Next
    Next
    o = ZeroField()
    For i = 0 To 14: t(i) = t(i) + 38# * t(i + 16): Next
    For i = 0 To 15: o(i) = t(i): Next
    Carry o: Carry o
    FMul = o
End Function
Private Function FSquare(ByVal a As Variant) As Variant
    FSquare = FMul(a, a)
End Function
Private Sub SwapField(ByRef a As Variant, ByRef b As Variant, ByVal bit As Long)
    Dim t As Variant
    If bit <> 0 Then
        t = a: a = b: b = t
    End If
End Sub
Private Function FInverse(ByVal a As Variant) As Variant
    Dim c As Variant, i As Long
    c = a
    For i = 253 To 0 Step -1
        c = FSquare(c)
        If i <> 2 And i <> 4 Then c = FMul(c, a)
    Next
    FInverse = c
End Function
Private Function PackField(ByVal a As Variant) As Byte()
    Dim t As Variant, m As Variant, i As Long, j As Long, borrow As Long
    Dim out(31) As Byte, v As Long
    t = a: Carry t: Carry t: Carry t
    For j = 0 To 1
        m = ZeroField()
        m(0) = t(0) - 65517#
        For i = 1 To 14
            borrow = 0: If m(i - 1) < 0 Then borrow = 1
            m(i) = t(i) - 65535# - borrow
            m(i - 1) = m(i - 1) - Int(m(i - 1) / 65536#) * 65536#
        Next
        borrow = 0: If m(14) < 0 Then borrow = 1
        m(15) = t(15) - 32767# - borrow
        borrow = 0: If m(15) < 0 Then borrow = 1
        m(14) = m(14) - Int(m(14) / 65536#) * 65536#
        SwapField t, m, 1 - borrow
    Next
    For i = 0 To 15
        v = CLng(t(i))
        out(2 * i) = v And 255
        out(2 * i + 1) = (v \ 256) And 255
    Next
    PackField = out
End Function
Public Function CVX25519(ByRef scalar() As Byte, ByRef point() As Byte) As Byte()
    Dim z() As Byte, x As Variant, a As Variant, b As Variant, c As Variant
    Dim d As Variant, e As Variant, f As Variant, k As Variant
    Dim i As Long, bit As Long
    If UBound(scalar) <> 31 Or UBound(point) <> 31 Then Err.Raise 5, , "X25519 requires 32 bytes"
    z = scalar: z(31) = (z(31) And 127) Or 64: z(0) = z(0) And 248
    x = ZeroField()
    For i = 0 To 15: x(i) = CDbl(point(2 * i)) + 256# * point(2 * i + 1): Next
    x(15) = CLng(x(15)) And 32767
    a = ZeroField(): a(0) = 1
    b = x: c = ZeroField(): d = ZeroField(): d(0) = 1
    k = ZeroField(): k(0) = 56129: k(1) = 1
    For i = 254 To 0 Step -1
        bit = (z(i \ 8) \ (2 ^ (i And 7))) And 1
        SwapField a, b, bit: SwapField c, d, bit
        e = FAdd(a, c): a = FSub(a, c)
        c = FAdd(b, d): b = FSub(b, d)
        d = FSquare(e): f = FSquare(a)
        a = FMul(c, a): c = FMul(b, e)
        e = FAdd(a, c): a = FSub(a, c)
        b = FSquare(a): c = FSub(d, f)
        a = FMul(c, k): a = FAdd(a, d)
        c = FMul(c, a): a = FMul(d, f)
        d = FMul(b, x): b = FSquare(e)
        SwapField a, b, bit: SwapField c, d, bit
    Next
    a = FMul(a, FInverse(c))
    CVX25519 = PackField(a)
End Function

' Demo only: 87-bit safe-prime finite-field DH. Decimal preserves exact integers.
Private Function DemoP() As Variant
    DemoP = CDec("154742504910672534362390399")
End Function
Private Function DemoQ() As Variant
    DemoQ = CDec("77371252455336267181195199")
End Function
Private Function ModD(ByVal n As Variant, ByVal p As Variant) As Variant
    ' Decimal division can round very near integer boundaries: adjust remainder.
    Dim q As Variant, r As Variant
    q = Fix(n / p): r = n - q * p
    If r < 0 Then r = r + p
    If r >= p Then r = r - p
    ModD = r
End Function
Private Function DemoMul(ByVal a As Variant, ByVal b As Variant) As Variant
    Dim r As Variant, p As Variant
    r = CDec(0): p = DemoP()
    Do While b > 0
        If ModD(b, CDec(2)) = 1 Then r = ModD(r + a, p)
        a = ModD(a + a, p): b = Fix(b / CDec(2))
    Loop
    DemoMul = r
End Function
Private Function DemoPow(ByVal a As Variant, ByVal n As Variant) As Variant
    Dim r As Variant
    r = CDec(1)
    Do While n > 0
        If ModD(n, CDec(2)) = 1 Then r = DemoMul(r, a)
        a = DemoMul(a, a): n = Fix(n / CDec(2))
    Loop
    DemoPow = r
End Function
Private Function DemoNumber(ByRef b() As Byte) As Variant
    Dim n As Variant, i As Long
    n = CDec(0)
    For i = 10 To 0 Step -1: n = n * CDec(256) + CDec(b(i)): Next
    DemoNumber = n
End Function
Private Function DemoBytes(ByVal n As Variant) As Byte()
    Dim b(10) As Byte, i As Long, r As Variant
    For i = 0 To 10
        r = ModD(n, CDec(256)): b(i) = CByte(r): n = (n - r) / CDec(256)
    Next
    DemoBytes = b
End Function
Public Function CVDemoPublic(ByRef privateKey() As Byte) As Byte()
    CVDemoPublic = DemoBytes(DemoPow(CDec(4), DemoNumber(privateKey)))
End Function
Public Function CVDemoShared(ByRef privateKey() As Byte, ByRef publicKey() As Byte) As Byte()
    Dim y As Variant, n As Variant
    y = DemoNumber(publicKey): n = DemoNumber(privateKey)
    If n < 2 Or n >= DemoQ() Then Err.Raise 5, , "Invalid demo private key"
    If y < 2 Or y >= DemoP() - 1 Then Err.Raise 5, , "Invalid demo public key"
    If DemoPow(y, DemoQ()) <> 1 Then Err.Raise 5, , "Demo key outside subgroup"
    CVDemoShared = DemoBytes(DemoPow(y, n))
End Function
Public Function CVDemoPrivate() As Byte()
    Dim b() As Byte, n As Variant
    Do
        b = CVRandom(11): b(10) = b(10) And 63
        n = DemoNumber(b)
    Loop While n < 2 Or n >= DemoQ()
    CVDemoPrivate = b
End Function
