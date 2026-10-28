$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path $PSScriptRoot -Parent
$securityPath = 'HKCU:/Software/Microsoft/Office/16.0/Excel/Security'
$oldValue = Get-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -ErrorAction SilentlyContinue
$hadValue = $null -ne $oldValue
$excel = $null
$book = $null
try {
    $null = New-Item -Path $securityPath -Force
    $null = New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -Value 1 -PropertyType DWord -Force
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $testDir = Join-Path $projectRoot '.test-output'
    $null = New-Item -ItemType Directory -Path $testDir -Force
    foreach ($mode in @('production','demo')) {
        $book = $excel.Workbooks.Add()
        foreach ($module in @('cellveilxl.bas','cv_tests.bas')) {
            $moduleRoot = $PSScriptRoot
            if ($module -eq 'cellveilxl.bas') { $moduleRoot = $projectRoot }
            $component = $book.VBProject.VBComponents.Import((Join-Path $moduleRoot $module))
            if ($module -eq 'cellveilxl.bas') {
                $code = $component.CodeModule.Lines(1,$component.CodeModule.CountOfLines)
                $code += [Environment]::NewLine + [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot 'carrier_tests.vba'))
                $code = $code.Replace('CRYPTO_MODE As String = "production"',('CRYPTO_MODE As String = "'+$mode+'"'))
                $code = $code.Replace('    MsgBox gStep & ": " & msg, vbCritical, "CellVeilXL"','    Err.Raise 5, , gStep & ": " & msg')
                # Automate only UI: execute the original macro bodies without blocking dialogs.
                $code = [regex]::Replace($code,'(?ms)^    MsgBox .*?(?=^    Exit Sub|^End Sub)','    Debug.Print "Test notification"' + "`r`n")
                $carrierPath = Join-Path $testDir ('example-'+$mode+'.xlsx')
                $code = $code.Replace('selected = Application.GetOpenFilename("Excel document (*.xlsx),*.xlsx")',('selected = "'+$carrierPath+'"'))
                $component.CodeModule.DeleteLines(1,$component.CodeModule.CountOfLines)
                $component.CodeModule.AddFromString($code)
            }
        }
        if ($mode -eq 'production') {
            $result = $excel.Run("'" + $book.Name + "'!CVCryptoSelfTest")
            Write-Output $result
            if (-not $result.StartsWith('PASS:')) { throw $result }
        }
        $result = $excel.Run("'" + $book.Name + "'!CVCarrierSelfTest", $testDir)
        Write-Output ($mode+': '+$result)
        if (-not $result.StartsWith('PASS:')) { throw $result }
        $source = $book.Worksheets.Item('List1')
        $source.Range('A1:A10').ClearContents()
        $source.Range('A1').Value2 = 'Confidential range test'
        $source.Range('A2').Value2 = 123.5
        $source.Range('A3').Value2 = $true
        $env:PYTHONPATH = Join-Path $projectRoot 'python'
        $keys = (& python -c "import json; from keygen import generate; print(json.dumps(generate('$mode')))" | ConvertFrom-Json)
        $source.Range('B1').Value2 = $keys[0]
        $source.Range('B2').Value2 = $keys[1]
        $hostSheet = $book.Worksheets.Add()
        $hostSheet.Name = 'List2'
        $hostSheet.Range('A1:C1').Merge()
        $hostSheet.Range('A1').Value2 = 'Operations report'
        $hostSheet.Range('A1').Font.Size = 18
        $hostSheet.Range('A1').Font.Bold = $true
        $hostSheet.Range('A2').Value2 = 'Section'
        $hostSheet.Range('B2').Value2 = 'Notes'
        $hostSheet.Range('C2').Value2 = 'Status'
        for ($row=3; $row -le 50; $row++) {
            $hostSheet.Cells.Item($row,1).Value2 = 'Review item '+($row-2)
            $hostSheet.Cells.Item($row,2).Value2 = 'Routine review completed; records prepared for the next reporting period.'
            $hostSheet.Cells.Item($row,3).Value2 = 'Complete'
        }
        $hostSheet.Range('A2:C50').Font.Size = 11
        $hostSheet.Range('A:A').ColumnWidth = 24
        $hostSheet.Range('B:B').ColumnWidth = 65
        $hostSheet.Range('C:C').ColumnWidth = 18
        $hostSheet.Range('A2:C50').WrapText = $true
        $hostSheet.Range('A3:C50').RowHeight = 34
        $hostSheet.Range('A2:C2').Font.Bold = $true
        $hostSheet.Range('A2:C2').Interior.Color = 15132390
        Write-Output "STEP EMBED"
        $null = $excel.Run("'" + $book.Name + "'!EmbedScr")
        Write-Output "STEP GET OUTPUT"
        $outputBook = $excel.ActiveWorkbook
        if ($outputBook.Name -eq $book.Name -or $outputBook.Worksheets.Count -ne 1 -or $outputBook.HasVBProject) { throw 'EmbedScr did not create a clean document' }
        Write-Output "STEP SAVE"
        $null = $excel.Run("'" + $book.Name + "'!CVSaveActiveTestDocument", [string]$carrierPath)
        $outputBook.Worksheets.Item(1).PageSetup.PrintArea = '$A$1:$C$15'
        $outputBook.Worksheets.Item(1).PageSetup.Orientation = 2
        $outputBook.Worksheets.Item(1).PageSetup.Zoom = $false
        $outputBook.Worksheets.Item(1).PageSetup.FitToPagesWide = 1
        Write-Output "STEP PDF"
        $null = $excel.Run("'" + $book.Name + "'!CVExportActiveTestDocument", (Join-Path $testDir ('example-'+$mode+'.pdf')))
        $outputBook.Close($false)
        Write-Output "STEP EXTRACT"
        $null = $excel.Run("'" + $book.Name + "'!ExtractScr")
        if ($source.Range('D1').Value2 -ne 'Confidential range test' -or $source.Range('D2').Value2 -ne 123.5 -or $source.Range('D3').Value2 -ne $true) { throw 'Full macro round trip failed' }
        Write-Output ($mode+': PASS: EmbedScr/ExtractScr complete macro bodies; carrier-only export')
        $book.Close($false); $book=$null
    }
} finally {
    try { if ($null -ne $book) { $book.Close($false) } } finally {
        try { if ($null -ne $excel) { $excel.Quit() } } finally {
        if ($hadValue) { Set-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -Value $oldValue.AccessVBOM } else { Remove-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -ErrorAction SilentlyContinue }
        }
    }
}
