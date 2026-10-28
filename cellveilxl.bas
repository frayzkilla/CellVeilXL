Option Explicit

Private Const SCR_SHEET As String = "List1"
Private Const HOST_SHEET As String = "List2"
Private Const SCR_CELL As String = "A1"
Private Const PWRD_CELL As String = "B1"
Private Const OUTPUT_CELL As String = "C1"
Private Const HOST_RANGE As String = ""

Private gStep As String
Private gCellAddr As String

Private Function GetColors() As Long()
    Dim c(0 To 15) As Long, i As Integer
    For i = 0 To 15
        c(i) = RGB(i, i, i)
    Next i
    GetColors = c
End Function

Private Function GetRotations() As Integer()
    Dim r(0 To 3) As Integer
    r(0) = 0: r(1) = 1: r(2) = 2: r(3) = 3
    GetRotations = r
End Function

Private Function CellIsUsable(cell As Range) As Boolean
    CellIsUsable = False
    If cell.HasFormula Then Exit Function
    If cell.MergeCells Then Exit Function
    If Len(CStr(cell.Value)) = 0 Then Exit Function
    CellIsUsable = True
End Function

Private Function ReadAngle(cell As Range) As Integer
    Dim v As Long
    v = CLng(cell.Orientation)
    Select Case v
        Case -4128: ReadAngle = 0
        Case -4166: ReadAngle = 255
        Case -4170: ReadAngle = 90
        Case -4171: ReadAngle = -90
        Case Else: ReadAngle = CInt(v)
    End Select
End Function

Private Function RC4Transform(data() As Byte, key As String) As Byte()
    Dim S(0 To 255) As Integer
    Dim i As Integer, j As Integer, t As Integer
    Dim kb() As Byte
    kb = StrConv(key, vbFromUnicode)
    Dim klen As Integer
    klen = UBound(kb) - LBound(kb) + 1
    If klen = 0 Then klen = 1
    For i = 0 To 255: S(i) = i: Next i
    j = 0
    For i = 0 To 255
        j = (j + S(i) + kb(LBound(kb) + (i Mod klen))) Mod 256
        t = S(i): S(i) = S(j): S(j) = t
    Next i
    Dim out() As Byte
    ReDim out(LBound(data) To UBound(data))
    i = 0: j = 0
    Dim k As Long
    For k = LBound(data) To UBound(data)
        i = (i + 1) Mod 256
        j = (j + S(i)) Mod 256
        t = S(i): S(i) = S(j): S(j) = t
        out(k) = data(k) Xor CByte(S((S(i) + S(j)) Mod 256))
    Next k
    RC4Transform = out
End Function

Private Function BytesToBits(data() As Byte) As Byte()
    Dim n As Long, bits() As Byte, k As Long, i As Long, b As Integer
    n = (UBound(data) - LBound(data) + 1) * 8
    ReDim bits(0 To n - 1)
    k = 0
    For i = LBound(data) To UBound(data)
        For b = 7 To 0 Step -1
            bits(k) = (data(i) \ (2 ^ b)) And 1
            k = k + 1
        Next b
    Next i
    BytesToBits = bits
End Function

Private Function ReadByteAt(bits() As Byte, startIdx As Long) As Long
    Dim v As Long, b As Integer
    v = 0
    For b = 0 To 7
        v = v * 2 + bits(startIdx + b)
    Next b
    ReadByteAt = v
End Function

Private Function FindIdxInt(arr() As Integer, val As Integer) As Integer
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        If arr(i) = val Then FindIdxInt = i: Exit Function
    Next i
    FindIdxInt = -1
End Function

Private Function FindIdxLong(arr() As Long, val As Long) As Integer
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        If arr(i) = val Then FindIdxLong = i: Exit Function
    Next i
    FindIdxLong = -1
End Function

Private Function BuildErrorMessage(ByVal prefix As String) As String
    Dim msg As String
    msg = prefix & vbCrLf & vbCrLf
    msg = msg & "Step: " & gStep & vbCrLf
    If Len(gCellAddr) > 0 Then msg = msg & "Cell: " & gCellAddr & vbCrLf
    msg = msg & "Err.Number: " & Err.Number & vbCrLf
    msg = msg & "Err.Description: " & Err.Description & vbCrLf & vbCrLf
    msg = msg & "Possible causes:" & vbCrLf
    msg = msg & "- Sheet names must be exactly 'List1' and 'List2'." & vbCrLf
    msg = msg & "- List2 must contain plain text, not formulas." & vbCrLf
    msg = msg & "- List2 must not be protected." & vbCrLf
    msg = msg & "- Do not edit or reformat List2 between embed and extract."
    BuildErrorMessage = msg
End Function

Sub EmbedScr()
    Dim wsS As Worksheet, wsH As Worksheet
    Dim scrText As String, pwrd As String
    Dim data() As Byte, enc() As Byte
    Dim dataLen As Long, i As Long
    Dim checksum(0 To 3) As Byte
    Dim packet() As Byte
    Dim bits() As Byte
    Dim totalBits As Long
    Dim target As Range
    Dim totalChars As Long, totalCells As Long
    Dim capacity As Long
    Dim cell As Range
    Dim colors() As Long, rots() As Integer
    Dim bitIdx As Long, charIdx As Long
    Dim rv As Integer, cv As Integer
    Dim txt As String
    Dim prevScreen As Boolean, prevCalc As XlCalculation, prevEvents As Boolean

    gStep = "Initialization"
    gCellAddr = ""
    On Error GoTo Fail

    Set wsS = ThisWorkbook.Sheets(SCR_SHEET)
    Set wsH = ThisWorkbook.Sheets(HOST_SHEET)

    gStep = "Read source text and password"
    scrText = CStr(wsS.Range(SCR_CELL).Value)
    pwrd = CStr(wsS.Range(PWRD_CELL).Value)
    If Len(scrText) = 0 Then
        MsgBox "No source text in " & SCR_SHEET & "!" & SCR_CELL, vbCritical, "EmbedScr"
        Exit Sub
    End If
    If Len(pwrd) = 0 Then
        MsgBox "Password is empty in " & SCR_SHEET & "!" & PWRD_CELL, vbCritical, "EmbedScr"
        Exit Sub
    End If

    If wsH.ProtectContents Then
        MsgBox "Sheet '" & HOST_SHEET & "' is protected. Unprotect it and try again.", _
               vbCritical, "EmbedScr"
        Exit Sub
    End If

    gStep = "Encrypt text with RC4"
    data = StrConv(scrText, vbFromUnicode)
    enc = RC4Transform(data, pwrd)
    dataLen = UBound(enc) - LBound(enc) + 1

    gStep = "Build packet: length + cipher + checksum"
    For i = 0 To dataLen - 1
        checksum(i Mod 4) = checksum(i Mod 4) Xor enc(LBound(enc) + i)
    Next i

    ReDim packet(0 To 4 + dataLen + 4 - 1)
    packet(0) = dataLen And 255
    packet(1) = (dataLen \ 256) And 255
    packet(2) = (dataLen \ 65536) And 255
    packet(3) = (dataLen \ 16777216) And 255
    For i = 0 To dataLen - 1
        packet(4 + i) = enc(LBound(enc) + i)
    Next i
    packet(4 + dataLen + 0) = checksum(0)
    packet(4 + dataLen + 1) = checksum(1)
    packet(4 + dataLen + 2) = checksum(2)
    packet(4 + dataLen + 3) = checksum(3)

    gStep = "Convert packet to bits"
    bits = BytesToBits(packet)
    totalBits = UBound(bits) - LBound(bits) + 1

    gStep = "Select carrier range on " & HOST_SHEET
    If Len(HOST_RANGE) > 0 Then
        Set target = wsH.Range(HOST_RANGE)
    Else
        Set target = wsH.UsedRange
    End If

    gStep = "Measure carrier capacity"
    For Each cell In target.Cells
        If CellIsUsable(cell) Then
            totalCells = totalCells + 1
            totalChars = totalChars + Len(CStr(cell.Value))
        End If
    Next cell
    capacity = totalChars * 4 + totalCells * 2
    If capacity < totalBits Then
        MsgBox "Not enough space in carrier!" & vbCrLf & vbCrLf & _
               "Bits needed: " & totalBits & vbCrLf & _
               "Bits available: " & capacity & vbCrLf & _
               "Usable cells (non-empty, non-merged, non-formula): " & totalCells & vbCrLf & _
               "Usable characters: " & totalChars & vbCrLf & vbCrLf & _
               "Unmerge header cells or add more plain-text cells to " & HOST_SHEET & ".", _
               vbCritical, "EmbedScr"
        Exit Sub
    End If

    gStep = "Disable screen updates / calculation / events"
    prevScreen = Application.ScreenUpdating
    prevCalc = Application.Calculation
    prevEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    colors = GetColors(): rots = GetRotations()
    bitIdx = 0

    gStep = "Write bits into cell formatting"
    For Each cell In target.Cells
        gCellAddr = cell.Address(False, False)
        If Not CellIsUsable(cell) Then GoTo NextCell
        txt = CStr(cell.Value)
        If VarType(cell.Value) <> vbString Then
            cell.NumberFormat = "@"
            cell.Value = txt
        End If
        If bitIdx + 2 <= totalBits Then
            rv = bits(bitIdx) * 2 + bits(bitIdx + 1)
            bitIdx = bitIdx + 2
            cell.Orientation = rots(rv)
        End If
        For charIdx = 1 To Len(txt)
            If bitIdx + 4 <= totalBits Then
                cv = bits(bitIdx) * 8 + bits(bitIdx + 1) * 4 + bits(bitIdx + 2) * 2 + bits(bitIdx + 3)
                bitIdx = bitIdx + 4
                cell.Characters(charIdx, 1).Font.Color = colors(cv)
            End If
        Next charIdx
NextCell:
    Next cell
    gCellAddr = ""

    gStep = "Restore Excel settings"
    Application.ScreenUpdating = prevScreen
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents

    If bitIdx < totalBits Then
        MsgBox "Only " & bitIdx & " of " & totalBits & " bits were embedded." & vbCrLf & vbCrLf & _
               "Some carrier cells were skipped (merged, formulas, empty)." & vbCrLf & _
               "Extraction will fail. Add more plain-text cells to " & HOST_SHEET & ".", _
               vbCritical, "EmbedScr"
        Exit Sub
    End If

    MsgBox "Done. Embedded " & bitIdx & " of " & totalBits & " bits." & vbCrLf & _
           "Save as .xlsx to strip the VBA module.", vbInformation, "EmbedScr"
    Exit Sub

Fail:
    gStep = gStep & " [FAILED]"
    On Error Resume Next
    Application.ScreenUpdating = prevScreen
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    On Error GoTo 0
    MsgBox BuildErrorMessage("EmbedScr failed:"), vbCritical, "EmbedScr"
End Sub

Sub ExtractScr()
    Dim wsS As Worksheet, wsH As Worksheet
    Dim pwrd As String
    Dim target As Range
    Dim colors() As Long, rots() As Integer
    Dim bits() As Byte
    Dim bitCount As Long, i As Long
    Dim cell As Range
    Dim rv As Integer, cv As Integer
    Dim charIdx As Long
    Dim txt As String
    Dim b0 As Long, b1 As Long, b2 As Long, b3 As Long
    Dim dataLen As Long
    Dim enc() As Byte
    Dim storedCs(0 To 3) As Byte, actualCs(0 To 3) As Byte
    Dim base As Long
    Dim decrypted() As Byte
    Dim result As String

    gStep = "Initialization"
    gCellAddr = ""
    On Error GoTo Fail

    Set wsS = ThisWorkbook.Sheets(SCR_SHEET)
    Set wsH = ThisWorkbook.Sheets(HOST_SHEET)

    gStep = "Read password"
    pwrd = CStr(wsS.Range(PWRD_CELL).Value)
    If Len(pwrd) = 0 Then
        MsgBox "Password is empty in " & SCR_SHEET & "!" & PWRD_CELL, vbCritical, "ExtractScr"
        Exit Sub
    End If

    gStep = "Select carrier range on " & HOST_SHEET
    If Len(HOST_RANGE) > 0 Then
        Set target = wsH.Range(HOST_RANGE)
    Else
        Set target = wsH.UsedRange
    End If

    colors = GetColors(): rots = GetRotations()
    ReDim bits(0 To 10000000)
    bitCount = 0

    gStep = "Read bits from cell formatting"
    For Each cell In target.Cells
        gCellAddr = cell.Address(False, False)
        If Not CellIsUsable(cell) Then GoTo NextCell
        txt = CStr(cell.Value)
        If VarType(cell.Value) <> vbString Then
            cell.NumberFormat = "@"
            cell.Value = txt
        End If
        rv = FindIdxInt(rots, ReadAngle(cell))
        If rv >= 0 Then
            bits(bitCount) = (rv \ 2) And 1
            bits(bitCount + 1) = rv And 1
        Else
            bits(bitCount) = 0: bits(bitCount + 1) = 0
        End If
        bitCount = bitCount + 2
        For charIdx = 1 To Len(txt)
            cv = FindIdxLong(colors, cell.Characters(charIdx, 1).Font.Color)
            If cv >= 0 Then
                bits(bitCount + 0) = (cv \ 8) And 1
                bits(bitCount + 1) = (cv \ 4) And 1
                bits(bitCount + 2) = (cv \ 2) And 1
                bits(bitCount + 3) = cv And 1
            Else
                bits(bitCount + 0) = 0: bits(bitCount + 1) = 0
                bits(bitCount + 2) = 0: bits(bitCount + 3) = 0
            End If
            bitCount = bitCount + 4
        Next charIdx
NextCell:
    Next cell
    gCellAddr = ""

    If bitCount < 64 Then
        MsgBox "Not enough bits collected: " & bitCount & "." & vbCrLf & vbCrLf & _
               "Possible causes:" & vbCrLf & _
               "- List2 is empty or too small." & vbCrLf & _
               "- Formatting on List2 was modified after embedding." & vbCrLf & _
               "- Wrong sheet names (must be List1 / List2).", _
               vbCritical, "ExtractScr"
        Exit Sub
    End If

    gStep = "Read packet length"
    b0 = ReadByteAt(bits, 0): b1 = ReadByteAt(bits, 8)
    b2 = ReadByteAt(bits, 16): b3 = ReadByteAt(bits, 24)
    dataLen = b0 + b1 * 256 + b2 * 65536 + b3 * 16777216

    If dataLen <= 0 Or dataLen > 5000000 Then
        MsgBox "Invalid packet length: " & dataLen & "." & vbCrLf & vbCrLf & _
               "The hidden data was probably damaged, or List2 does not contain" & vbCrLf & _
               "a steganographic payload (e.g. wrong sheet).", _
               vbCritical, "ExtractScr"
        Exit Sub
    End If
    If bitCount < 32 + dataLen * 8 + 32 Then
        MsgBox "Not enough bits for the declared payload." & vbCrLf & vbCrLf & _
               "Needed bits: " & (32 + dataLen * 8 + 32) & vbCrLf & _
               "Collected bits: " & bitCount & vbCrLf & _
               "Declared length: " & dataLen & " bytes", _
               vbCritical, "ExtractScr"
        Exit Sub
    End If

    gStep = "Read cipher bytes"
    ReDim enc(0 To dataLen - 1)
    For i = 0 To dataLen - 1
        enc(i) = CByte(ReadByteAt(bits, 32 + i * 8))
    Next i

    gStep = "Verify checksum"
    base = 32 + dataLen * 8
    storedCs(0) = CByte(ReadByteAt(bits, base + 0))
    storedCs(1) = CByte(ReadByteAt(bits, base + 8))
    storedCs(2) = CByte(ReadByteAt(bits, base + 16))
    storedCs(3) = CByte(ReadByteAt(bits, base + 24))
    For i = 0 To dataLen - 1
        actualCs(i Mod 4) = actualCs(i Mod 4) Xor enc(i)
    Next i
    For i = 0 To 3
        If storedCs(i) <> actualCs(i) Then
            MsgBox "Integrity check failed." & vbCrLf & vbCrLf & _
                   "Stored checksum byte " & i & ": " & storedCs(i) & vbCrLf & _
                   "Computed checksum byte " & i & ": " & actualCs(i) & vbCrLf & vbCrLf & _
                   "Possible causes:" & vbCrLf & _
                   "- Formatting on List2 was changed after embedding." & vbCrLf & _
                   "- Wrong password." & vbCrLf & _
                   "- Wrong carrier range on List2.", _
                   vbCritical, "ExtractScr"
            Exit Sub
        End If
    Next i

    gStep = "Decrypt with RC4"
    decrypted = RC4Transform(enc, pwrd)
    result = StrConv(decrypted, vbUnicode)

    gStep = "Write result"
    wsS.Range(OUTPUT_CELL).Value = result
    MsgBox "Done. " & Len(result) & " characters written to " & _
           SCR_SHEET & "!" & OUTPUT_CELL, vbInformation, "ExtractScr"
    Exit Sub

Fail:
    gStep = gStep & " [FAILED]"
    MsgBox BuildErrorMessage("ExtractScr failed:"), vbCritical, "ExtractScr"
End Sub