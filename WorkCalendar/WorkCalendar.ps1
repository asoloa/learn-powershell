# StrictMode for development only
# Set-StrictMode -Version Latest

Add-Type -AssemblyName System.Windows.Forms

Function Get-InputFile {
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

Function Get-EmployeeData {
    param(
        [Parameter(Mandatory)] [string]$wsName,
        [Parameter(Mandatory)] [object]$workbook
    )

    $data_ws = $workbook.Sheets.Item($wsName)
    $lastRow = $data_ws.Cells($data_ws.Rows.Count, "G").End(3).Row  # 3 = xlUp
    $usedRng = $data_ws.Range("A1:G$lastRow")
    $arr = $usedRng.Value2

    # Dimensions (PowerShell arrays from Excel are 1-based, VB-style)
    $rowCount = $arr.GetUpperBound(0)
    $colCount = $arr.GetUpperBound(1)

    $rowData = New-Object string[] ($colCount)

    # Use a StringBuilder for performance
    $sb = New-Object System.Text.StringBuilder

    for ($row = 1; $row -le $rowCount; $row++) {
        for ($col = 1; $col -le $colCount; $col++) {
            $val = $arr[$row, $col]
            # If cell in a row is empty, the current row is a leave entry (see [Data Sheet])
            # For uniformity and ease of pattern-matching, we set the current element to the previous element's value
            if ($null -eq $val -or $val -eq "") {
                if ($col -gt 1) {
                    # Enclosing "$col - 1" in parentheses is necessary to properly evaluate the arithmetic operation
                    # Not doing so will result to "Method invocation failed because [System.Object[]] does not contain a method named 'op_Subtraction'." error.
                    $val = $arr[$row, ($col - 1)]
                }
                else {
                    $val = ""
                }
            }
            $arr[$row, $col] = $val
            $rowData[$col - 1] = $val
        }
        # Join row values with tabs, append newline
        # NOTE: .AppendLine to returns the StringBuilder object to the pipeline, PowerShell will format it for output.
        #   [void] prevents it from returning the SB object.
        #   Also, without [void], the receiving string may contain multiple instances of the SB values.
        [void]$sb.AppendLine(($rowData -join "`t"))
    }
    return $sb.ToString()
}

function Populate-Calendar {
    param(
        [Parameter(Mandatory)] [string]$dataString,
        [Parameter(Mandatory)] [int]$lastRow,
        [Parameter(Mandatory)] [int]$calendarDays,
        [Parameter(Mandatory)] [object]$workbook
    )

    $reportWs = $workbook.Sheets.Item("Calendar Sheet")
    $lines = $dataString -split "`r`n"

    # employeeCollection[emp][project][dayIndex] = "hours" / "VL" / "SL"
    $employeeCollection = @{}

    foreach ($line in $lines) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }

        $fields = $line -split "`t"
        if ($fields.Count -lt 7) { continue }

        $employeeName  = $fields[0]
        $dateStr  = $fields[2]
        $projectStr   = $fields[3]
        $hoursStr = $fields[5]
        $category = $fields[6]

        # Build dictionary chain
        if (-not $employeeCollection.ContainsKey($employeeName)) {
            $employeeCollection[$employeeName] = @{}
        }
        if (-not $employeeCollection[$employeeName].ContainsKey($projectStr)) {
            $employeeCollection[$employeeName][$projectStr] = @{}
        }

        # DayIndex = 0-based
        $dayIndex = ([DateTime]::FromOADate($dateStr)).Day - 1

        # Determine stored value
        if ($category -notlike "*Leave*") {
            # [string]::Format("{0:N2}", [double]$hoursStr) → Behaves the same way as below
            $valueStr = ("{0:N2}" -f [double]$hoursStr)
        }
        elseif ($category -like "*Vacation Leave*") {
            $valueStr = "VL"
        }
        elseif ($category -like "*Sick Leave*") {
            $valueStr = "SL"
        }
        else {
            $valueStr = "UNK"
        }

        # Store
        $employeeCollection[$employeeName][$projectStr][$dayIndex] = $valueStr
    }

    # Prebuild blank hours row
    $blankHours = @("") * $calendarDays

    # Count needed rows
    $totalRows = 0
    foreach ($emp in $employeeCollection.Keys) {
        $totalRows++               # employee header
        $totalRows += $employeeCollection[$emp].Count  # project rows
        $totalRows++               # spacer
    }

    # Allocate output array (1-based to match Excel)
    $outputArr = New-Object 'object[,]' $totalRows, ($calendarDays + 1)
    $rowType   = New-Object 'int[]' $totalRows  # 0 = spacer/blank, 1 = header, 2 = project row
    $rowOut = 0

    foreach ($emp in $employeeCollection.Keys) {
        # EMPLOYEE HEADER
        $outputArr[$rowOut, 0] = $emp
        $rowType[$rowOut] = 1
        $rowOut++

        foreach ($project in $employeeCollection[$emp].Keys) {
            $outputArr[$rowOut, 0] = $project
            $rowType[$rowOut] = 2

            # copy fresh blank hours
            $workHours = $blankHours.Clone()

            # Fill hours for each day
            for ($day = 0; $day -lt $calendarDays; $day++) {
                if ($employeeCollection[$emp][$project].ContainsKey($day)) {
                    $workHours[$day] = $employeeCollection[$emp][$project][$day]
                }
                else {
                    $workHours[$day] = ""
                }
            }

            # Copy into array
            for ($day = 0; $day -lt $calendarDays; $day++) {
                $outputArr[$rowOut, ($day + 1)] = $workHours[$day]
            }
            $rowOut++
        }
        # spacer row
        $rowOut++
    }

    $cellStart = $reportWs.Range("B4")
    $cellEnd   = $cellStart.Offset(($totalRows - 1), $calendarDays)

    $range = $cellStart.Resize($totalRows, ($calendarDays + 1))
    $range.Value2 = $outputArr
    $range.NumberFormat = "0.00"
    $range.HorizontalAlignment = -4108   # xlCenter
    $range.VerticalAlignment   = -4108   # xlCenter

    # Format rows based on type: 0 = spacer/blank, 1 = header, 2 = project row
    for ($row = 0; $row -lt $totalRows; $row++) {
        switch ($rowType[$row]) {
            1 {  # EMPLOYEE HEADER
                $cell = $cellStart.Offset($row, 0)
                $label = $outputArr[$row, 0]
                $cell.Value = (Get-Culture).TextInfo.ToTitleCase($label.ToLower())
                $cell.IndentLevel = 1
                $cell.Font.Bold = $true
                $cell.Resize(1, ($calendarDays + 1)).Interior.Color = 0x00FBEDCA
            }
            2 {  # PROJECT ROW
                $cellStart.Offset($row, 0).IndentLevel = 3
            }
        }
    }

    # Borders
    $reportWs.Range($cellStart, $cellEnd).Borders.LineStyle = 1  # xlContinuous

    # External helpers
    Set-LeaveFormatting -Worksheet $reportWs -LastCell $cellEnd
    Shade-Weekends -Worksheet $reportWs -CalendarDays $calendarDays
}

function Set-LeaveFormatting {
    param(
        [Parameter(Mandatory)] [object]$worksheet,
        [Parameter(Mandatory)] [object]$lastCell
    )

    # Target range: C4 : lastCell
    $rng = $worksheet.Range("C4", $lastCell.Address(0,0))

    # Add conditional formats
    $fc1 = $rng.FormatConditions.Add(1, 3, "VL")  # xlCellValue=1, xlEqual=3
    $fc2 = $rng.FormatConditions.Add(1, 3, "SL")
    $fc3 = $rng.FormatConditions.Add(1, 3, "UNK")

    # Set interior colors
    $fc1.Interior.Color = 0x00FFFF  # vbYellow
    $fc2.Interior.Color = 0xFF00FF  # vbMagenta
    $fc3.Interior.Color = 0xFFFF00  # vbCyan
}

# Convert a column index (1-based) to Excel letters
function Get-ColumnLetter([int]$colIndex) {
    $col = ""
    while ($colIndex -gt 0) {
        $remainder = ($colIndex - 1) % 26
        $col = [char]($remainder + 65) + $col
        $colIndex = [Math]::Floor(($colIndex - $remainder) / 26)
    }
    return $col
}

function Get-WeekendRange {
    param(
        [Parameter(Mandatory)] [object]$worksheet,
        [Parameter(Mandatory)] [object]$cellIterator,
        [Parameter(Mandatory)] [int]$lastRow
    )

    # Start cell: one row below the header
    $startCell = $cellIterator.Offset(1, 0).Address(0, 0)

    # Extract column letter from column index
    $colIndex  = $cellIterator.Column
    $colLetter = Get-ColumnLetter $colIndex

    # End cell: <ColumnLetter><lastRow>  e.g., "C45"
    $endCell = "$colLetter$lastRow"

    # Return Excel range address
    return "${startCell}:${endCell}"
}

function Shade-Weekends {
    param(
        [Parameter(Mandatory)] [object]$worksheet,
        [Parameter(Mandatory)] [int]$calendarDays
    )

    $cellIterator = $worksheet.Range("C3")

    # Find last row in column B
    $lastRow = $worksheet.Cells($worksheet.Rows.Count, "B").End(-4162).Row + 1  # xlUp = -4162

    while ($cellIterator.Value2 -ne $null -and $cellIterator.Value2 -ne "") {
        $dt = [DateTime]::FromOADate($cellIterator.Value2)
        $wkday = $dt.DayOfWeek  # Sunday=0, Monday=1, ..., Saturday=6

        # VBA weekend: Sunday(1) or Saturday(7)
        if ($wkday -eq "Sunday" -or $wkday -eq "Saturday") {
            $worksheet.Range($(Get-WeekendRange -Worksheet $worksheet -CellIterator $cellIterator -LastRow $lastRow)).Interior.Pattern = 13 # xlPatternLightDown
        }

        $cellIterator = $cellIterator.Offset(0,1)
    }
}

Function Generate-DataSheet {
    param (
        [Parameter(Mandatory)] [string]$excelInput
    )

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $inputFile = $excel.Workbooks.Open($excelInput)
    $outputFile = $excel.Workbooks.Add()

    $inDataSheet = $inputFile.Worksheets.Item(1)
    $outDataSheet = $outputFile.Worksheets.Item(1)
    $outDataSheet.Name = "Data Sheet"
    $destCell = $outDataSheet.Range("A1")

    $lastUsedRow = $inDataSheet.Cells.Item($excel.Rows.Count, 1).End(-4162).Row # xlUp = -4162

    foreach ($col in @("A", "B", "C", "E", "F", "G", "H")) {
        $inDataSheet.Range("$($col)2", "$($col)$($lastUsedRow)").Copy() | Out-Null
        $destCell.PasteSpecial(-4163) | Out-Null # xlPasteValues = -4163 
        $destCell = $destCell.Offset(0, 1)
    }

    $outDataSheet.Range("C:C").NumberFormat = "mm/dd/yyyy"
    $outDataSheet.Range("F:F").NumberFormat = "#0.000"
    $outDataSheet.Range("A:G").Columns.AutoFit() | Out-Null
    $outDataSheet.Range("A:G").HorizontalAlignment = -4108

    $dataString =  $(Get-EmployeeData -wsName $outDataSheet.Name -workbook $outputFile)

    ##### Generate Calendar Sheet
    $outCalendarSheet = $outputFile.Worksheets.Add()
    $outCalendarSheet.Name = "Calendar Sheet"
    $outCalendarSheet.Columns.ColumnWidth = 3
    $outCalendarSheet.Range("B1").ColumnWidth = 30
    
    # TODO: Handle overflowing dates (dates from data sheet include previous and/or next months)
    $latestDate = $excel.WorksheetFunction.Max($outDataSheet.Range("C:C"))
    $latestDate = [DateTime]::FromOADate($latestDate)
    $year  = $latestDate.Year
    $month = $latestDate.Month
    $calendarDays = [DateTime]::DaysInMonth($year, $month)

    # Create an array that holds the month's days
    # Excel needs a 2D array: 1 row × N columns
    $dates2D = New-Object 'object[,]' 1, $calendarDays
    for ($i = 0; $i -lt $calendarDays; $i++) {
        $dates2D[0, $i] = [DateTime]::New($year, $month, $i + 1)
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

    Populate-Calendar -CalendarDays $calendarDays -LastRow $lastUsedRow -DataString $dataString -workbook $outputFile

    $excel.ActiveWindow.DisplayGridlines = $false
    $outDataSheet.Protect() | Out-Null
    $outCalendarSheet.Protect() | Out-Null

    ##### Save and exit files. Cleanup.
    $outputDirectory = $PSScriptRoot + "\extracted"
    [IO.Directory]::CreateDirectory($outputDirectory) | Out-Null
    $outputFilename = $outputDirectory + "\$(Get-Date -Format "yyyyMMdd-HHmmss").xlsx"
    $outputFile.SaveAs($outputFilename)
    $outputFile.Close()
    $inputFile.Close()
    $excel.Quit()

    Write-Host "Calendar file generated at '$outputFileName'"
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}  

# Main Block
$inputFile = Get-InputFile
if ($inputFile) {
    Generate-DataSheet -ExcelInput $inputFile
}

Write-Host "Performing cleanup..."
[GC]::Collect()
[GC]::WaitForPendingFinalizers()