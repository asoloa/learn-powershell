Function Get-InputFile {
    # Load the System.Windows.Forms assembly
    Add-Type -AssemblyName System.Windows.Forms

    # Create an instance of the OpenFileDialog class
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog

    $openFileDialog.InitialDirectory = $PWD.Path
    $openFileDialog.Filter = "Excel files (*.xls)|*.xls"
    $openFileDialog.Title = "Select input file"

    # Show the dialog box and capture the result
    $result = $openFileDialog.ShowDialog()

    # Check if a file was selected
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $openFileDialog.FileName
    }
    return
}

Function Generate-DataSheet {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ExcelInput
    )

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $inputFile = $excel.Workbooks.Open($ExcelInput)
    $outputFile = $excel.Workbooks.Add()

    $inDataSheet = $inputFile.Worksheets.Item(1)
    $outDataSheet = $outputFile.Worksheets.Item(1)
    $outDataSheet.Name = "Data Sheet"
    $destCell = $outDataSheet.Range("A1")

    $lastUsedRow = $inDataSheet.Cells.Item($excel.Rows.Count, 1).End(-4162).Row # xlUp = -4162

    foreach ($col in @("A", "B", "C", "F", "G", "H", "L")) {
        $inDataSheet.Range("$($col)2", "$($col)$($lastUsedRow)").Copy() | Out-Null
        $destCell.PasteSpecial(-4163) # xlPasteValues = -4163
        $destCell = $destCell.Offset(0, 1)
    }

    $outDataSheet.Range("C:C").NumberFormat = "mm/dd/yyyy"
    $outDataSheet.Range("F:F").NumberFormat = "#0.000"
    $outDataSheet.Range("A:G").Columns.AutoFit()
    $outDataSheet.Range("A:G").HorizontalAlignment = -4108
    $outDataSheet.Protect | Out-Null

    ##### Generate Calendar Sheet
    $outCalendarSheet = $outputFile.Worksheets.Add()
    $outCalendarSheet.Name = "Calendar Sheet"
    $outCalendarSheet.Columns.ColumnWidth = 3
    $outCalendarSheet.Range("B1").ColumnWidth = 30
    
    # TODO: Handle overflowing dates (dates from data sheet include previous and/or next months)
    # $earliestDate = $excel.WorksheetFunction.Min($outDataSheet.Range("C:C"))
    $latestDate = $excel.WorksheetFunction.Max($outDataSheet.Range("C:C"))
    $latestDate = [DateTime]::FromOADate($latestDate)
    $year  = $latestDate.Year
    $month = $latestDate.Month
    $calendarDays = [DateTime]::DaysInMonth($year, $month)

    # Create an array that holds the month's days
    # Excel needs a 2D array: 1 row × N columns
    $dates2D = New-Object 'object[,]' 1, $calendarDays
    for ($i = 0; $i -lt $calendarDays; $i++) {
        $dates2D[0, $i] = [datetime]::new($year, $month, $i + 1)
    }

    # Set formatting of Month-Year header
    $MYHeaderRange = $outCalendarSheet.Range("C2").Resize(1, $calendarDays)
    $MYHeaderRange.Merge()
    $MYHeaderRange.NumberFormat = "@"
    $MYHeaderRange.Value = ([CultureInfo]::InvariantCulture).DateTimeFormat.GetMonthName($month).ToUpper() + " " + $year
    $MYHeaderRange.Interior.Color = 0x83A9F1
    $MYHeaderRange.Font.Bold = $true

    # Set common formatting of Month-Year header and Days' cells
    $MYHeaderRange = $MYHeaderRange.Resize(2, $calendarDays)
    $MYHeaderRange.ColumnWidth = 5
    $MYHeaderRange.HorizontalAlignment = -4108 # xlCenter
    $MYHeaderRange.VerticalAlignment = -4108 # xlCenter
    $MYHeaderRange.Borders.LineStyle = 1 # xlContinuous

    # Set formatting of Days' cells
    $daysRange = $outCalendarSheet.Range("C3").Resize(1, $calendarDays)
    $daysRange.NumberFormat = "dd"
    $daysRange.Value2 = $dates2D  
    $daysRange.Interior.Color = 0xD5E2FB

    $excel.ActiveWindow.DisplayGridlines = $false

    ##### Save and exit files. Cleanup.
    $outputFile.SaveAs($PWD.Path + "\extracted_$(Get-Date -Format "yyyyMMdd-HHmmss").xlsx")
    $outputFile.Close()
    $inputFile.Close()
    $excel.Quit()

    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($indataSheet) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outdataSheet) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outputFile) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($inputFile) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    # Remove-Variable excel, inputFile, outputFile, inDataSheet, outDataSheet, destCell, lastUsedRow
}

# Main Block
$inputFile = Get-InputFile
Generate-DataSheet -ExcelInput $inputFile
