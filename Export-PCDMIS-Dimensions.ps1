# Copyright (C) 2026 Townsend Precision Labs Ltd
# SPDX-License-Identifier: GPL-3.0-or-later

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$OutputPath = "",

    [Parameter(Position = 1)]
    [string]$Operator = "",

    [Parameter(Position = 2)]
    [string]$JobNumber = "",

    [Parameter(Position = 3)]
    [string]$SerialNumber = "",

    [Parameter(Position = 4)]
    [string]$PartNumber = "",

    [Parameter(Position = 5)]
    [string]$Revision = "",

    [Parameter(Position = 6)]
    [string]$ProgramPath = "",

    [Parameter(Position = 7)]
    [string]$BridgePath = "",

    [switch]$Mock
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

# PC-DMIS can publish its Automation object in the Running Object Table even
# when a damaged installation has lost the PCDLRN.Application ProgID registry
# entry. Calling OleAut32 GetActiveObject with the stable Application CLSID
# keeps the exporter working in that condition.
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class PCDMISActiveObjectNative {
    [DllImport("oleaut32.dll")]
    public static extern int GetActiveObject(
        ref Guid rclsid,
        IntPtr reserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object obj);
}
'@

function Get-ComProperty {
    param(
        [Parameter(Mandatory = $true)]$Object,
        [Parameter(Mandatory = $true)][string]$Name,
        $Default = $null
    )

    if ($null -eq $Object) { return $Default }
    try {
        $value = $Object.$Name
        if ($null -eq $value) { return $Default }
        return $value
    }
    catch {
        return $Default
    }
}

function Convert-ToDouble {
    param($Value)
    try {
        if ($null -eq $Value -or $Value -eq "") { return 0.0 }
        return [Convert]::ToDouble($Value, [Globalization.CultureInfo]::InvariantCulture)
    }
    catch {
        try { return [Convert]::ToDouble($Value) }
        catch { return 0.0 }
    }
}

function Release-ComObject {
    param($Object)
    if ($null -ne $Object) {
        try {
            if ([Runtime.InteropServices.Marshal]::IsComObject($Object)) {
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
            }
        }
        catch { }
    }
}

function Get-PCDMISApplication {
    try {
        return [Runtime.InteropServices.Marshal]::GetActiveObject("PCDLRN.Application")
    }
    catch {
        $progIdError = $_.Exception.Message
    }

    try {
        $pcDmisRoot = Get-ChildItem (Join-Path $env:ProgramFiles "Hexagon") -Directory -ErrorAction Stop |
            Where-Object { Test-Path (Join-Path $_.FullName "PcdmisAutomationHelper.dll") } |
            Sort-Object Name -Descending |
            Select-Object -First 1
        if ($null -eq $pcDmisRoot) { throw "PcdmisAutomationHelper.dll was not found." }

        $rootPath = $pcDmisRoot.FullName
        [void][Reflection.Assembly]::LoadFrom((Join-Path $rootPath "PcdmisAutomationHelper.dll"))
        $application = [PcdmisAutomationHelper.Services.ApplicationService]::ConnectToPcdmis()
        if ($null -eq $application) { throw "ApplicationService returned no application object." }
        return $application
    }
    catch {
        $helperError = $_.Exception.Message
    }

    try {
        $applicationClsid = [Guid]"1E7404B9-F9B7-46A9-97A8-8D3049E78A9A"
        $application = $null
        $result = [PCDMISActiveObjectNative]::GetActiveObject(
            [ref]$applicationClsid,
            [IntPtr]::Zero,
            [ref]$application)
        if ($result -ne 0 -or $null -eq $application) {
            throw "OleAut32 GetActiveObject returned HRESULT 0x$($result.ToString('X8'))."
        }
        return $application
    }
    catch {
        throw "No usable PC-DMIS Automation instance was found. ProgID lookup failed: $progIdError ApplicationService failed: $helperError CLSID lookup failed: $($_.Exception.Message)"
    }
}

function Get-PCDMISVariableText {
    param([Parameter(Mandatory = $true)][string]$Name)

    $application = $null
    $partProgram = $null
    $variable = $null
    try {
        $application = Get-PCDMISApplication
        $partProgram = Get-ComProperty $application "ActivePartProgram" $null
        if ($null -eq $partProgram) { return "" }

        try { $variable = $partProgram.GetVariableValue($Name) }
        catch { return "" }
        if ($null -eq $variable) { return "" }

        try {
            $stringValue = [string]$variable.StringValue
            if ($stringValue -ne "") { return $stringValue }
        }
        catch { }

        try { return [Convert]::ToString($variable.DoubleValue, [Globalization.CultureInfo]::InvariantCulture) }
        catch { }
        try { return [string]$variable.LongValue }
        catch { }
        return ""
    }
    finally {
        Release-ComObject $variable
        Release-ComObject $partProgram
        Release-ComObject $application
    }
}

function New-DimensionRow {
    param(
        [int]$Item,
        [string]$ReportLabel,
        [string]$DimensionId,
        [string]$Axis,
        [string]$Features,
        [double]$Nominal,
        [double]$PlusTolerance,
        [double]$MinusTolerance,
        [double]$Measured,
        [double]$Deviation,
        [double]$OutTolerance,
        [double]$Bonus,
        [double]$Minimum,
        [double]$Maximum,
        [int]$Execution,
        [int]$CommandIndex,
        [string]$CommandType
    )

    $status = "PASS"
    if ([Math]::Abs($OutTolerance) -gt 0.0000001) { $status = "FAIL" }

    [pscustomobject][ordered]@{
        Item             = $Item
        ReportLabel      = $ReportLabel
        Dimension        = $DimensionId
        Axis             = $Axis
        Features         = $Features
        Nominal          = $Nominal
        PlusTolerance    = $PlusTolerance
        MinusTolerance   = $MinusTolerance
        Measured         = $Measured
        Deviation        = $Deviation
        OutTolerance     = $OutTolerance
        Bonus            = $Bonus
        Minimum          = $Minimum
        Maximum          = $Maximum
        Status           = $status
        Execution        = $Execution
        CommandIndex     = $CommandIndex
        CommandType      = $CommandType
    }
}

function Get-MockDimensions {
    $rows = New-Object System.Collections.Generic.List[object]
    $rows.Add((New-DimensionRow 1 "Example diameter 10.00 +/- 0.10" "DIM_DIAMETER" "D" "EXAMPLE_CYLINDER" 10.000 0.100 0.100 10.025 0.025 0.000 0.000 10.025 10.025 1 1 "Mock"))
    $rows.Add((New-DimensionRow 2 "Example length 25.00 +/- 0.20" "DIM_LENGTH" "M" "EXAMPLE_POINT_1; EXAMPLE_POINT_2" 25.000 0.200 0.200 25.230 0.230 0.030 0.000 25.230 25.230 1 2 "Mock"))
    $rows.Add((New-DimensionRow 3 "Example diameter 6.00 +/- 0.05" "DIM_BORE" "D" "EXAMPLE_BORE" 6.000 0.050 0.050 6.018 0.018 0.000 0.000 6.018 6.018 1 3 "Mock"))
    return $rows
}

function Get-PCDMISBridgeDimensions {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "PC-DMIS BASIC bridge file was not found: $Path"
    }

    $bridgeEncoding = [Text.Encoding]::GetEncoding(1252)
    $bridgeLines = [IO.File]::ReadAllLines($Path, $bridgeEncoding)
    $records = @($bridgeLines | ConvertFrom-Csv -Delimiter "`t")
    if ($records.Count -eq 0) {
        throw "PC-DMIS BASIC bridge file contained no dimension rows: $Path"
    }

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($record in $records) {
        $rows.Add((New-DimensionRow `
            -Item ([int](Convert-ToDouble $record.Item)) `
            -ReportLabel ([string]$record.ReportLabel) `
            -DimensionId ([string]$record.Dimension) `
            -Axis ([string]$record.Axis) `
            -Features ([string]$record.Features) `
            -Nominal (Convert-ToDouble $record.Nominal) `
            -PlusTolerance ([Math]::Abs((Convert-ToDouble $record.PlusTolerance))) `
            -MinusTolerance ([Math]::Abs((Convert-ToDouble $record.MinusTolerance))) `
            -Measured (Convert-ToDouble $record.Measured) `
            -Deviation (Convert-ToDouble $record.Deviation) `
            -OutTolerance (Convert-ToDouble $record.OutTolerance) `
            -Bonus (Convert-ToDouble $record.Bonus) `
            -Minimum (Convert-ToDouble $record.Minimum) `
            -Maximum (Convert-ToDouble $record.Maximum) `
            -Execution ([int](Convert-ToDouble $record.Execution)) `
            -CommandIndex ([int](Convert-ToDouble $record.CommandIndex)) `
            -CommandType ([string]$record.CommandType)
        ))
    }
    return $rows
}

function Get-PCDMISProgramDimensions {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "PC-DMIS program text file was not found: $Path"
    }

    try {
        $utf8 = New-Object Text.UTF8Encoding($false, $true)
        $programText = [IO.File]::ReadAllText($Path, $utf8)
    }
    catch {
        $programText = [IO.File]::ReadAllText($Path, [Text.Encoding]::Default)
    }
    $lines = @($programText -split "`r?`n")
    $rows = New-Object System.Collections.Generic.List[object]
    $item = 0
    $numberPattern = '[-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][-+]?\d+)?'

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = [string]$lines[$i]
        if ($line -notmatch '^\s*DIM\s+([^=]+)=\s*(.*?)\s+UNITS\s*=') { continue }

        $dimensionId = $matches[1].Trim()
        $description = $matches[2].Trim()
        $reportLabel = $description
        for ($back = $i - 1; $back -ge [Math]::Max(0, $i - 12); $back--) {
            $candidate = ([string]$lines[$back]).Trim()
            if ($candidate -match '^\d+[A-Za-z]?\)?\s*[-&)]') {
                $reportLabel = $candidate
                break
            }
        }

        $headerIndex = -1
        for ($j = $i + 1; $j -lt [Math]::Min($lines.Count, $i + 20); $j++) {
            if ([string]$lines[$j] -match '^\s*AX\s+NOMINAL\s+\+TOL\s+-TOL\s+MEAS\s+DEV\s+OUTTOL') {
                $headerIndex = $j
                break
            }
        }
        if ($headerIndex -lt 0) { continue }

        for ($j = $headerIndex + 1; $j -lt [Math]::Min($lines.Count, $headerIndex + 40); $j++) {
            $resultLine = [string]$lines[$j]
            $pattern = '^\s*(\S+)\s+(' + $numberPattern + ')\s+(' + $numberPattern + ')\s+(' + $numberPattern + ')\s+(' + $numberPattern + ')\s+(' + $numberPattern + ')\s+(' + $numberPattern + ')(?:\s|$)'
            if ($resultLine -notmatch $pattern) {
                if ($j -gt ($headerIndex + 1)) { break }
                continue
            }

            $item++
            $axis = $matches[1]
            $nominal = Convert-ToDouble $matches[2]
            $plusTol = [Math]::Abs((Convert-ToDouble $matches[3]))
            $minusTol = [Math]::Abs((Convert-ToDouble $matches[4]))
            $measured = Convert-ToDouble $matches[5]
            $deviation = Convert-ToDouble $matches[6]
            $outTol = Convert-ToDouble $matches[7]
            $rows.Add((New-DimensionRow $item $reportLabel $dimensionId $axis $description `
                $nominal $plusTol $minusTol $measured $deviation $outTol 0.0 `
                $measured $measured 1 ($i + 1) "ProgramText"))
        }
    }

    # ISO SIZE commands can contain a measured upper and lower local size but
    # no legacy AX result row. Preserve those as one summary row.
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = [string]$lines[$i]
        if ($line -notmatch '^\s*([^=\s]+)\s*=\s*SIZE/') { continue }
        $dimensionId = $matches[1].Trim()
        $block = New-Object System.Collections.Generic.List[string]
        $block.Add($line)
        for ($j = $i + 1; $j -lt [Math]::Min($lines.Count, $i + 25); $j++) {
            $candidate = [string]$lines[$j]
            if ($candidate -match '^\S.*=' -or $candidate -match '^\s{0,2}DIM\s+') { break }
            $block.Add($candidate)
        }
        $text = $block -join "`n"
        if ($text -notmatch '(?im)^\s*NOMINAL SIZE:\s*(' + $numberPattern + ')') { continue }
        $nominal = Convert-ToDouble $matches[1]
        $plusTol = 0.0; $minusTol = 0.0; $upperSize = $nominal; $lowerSize = $nominal; $feature = ""
        $hasUpperSize = $false; $hasLowerSize = $false; $hasUame = $false; $hasLocalSize = $false
        $uame = 0.0; $localSize = 0.0
        if ($text -match '(?im)^\s*UPPER TOLERANCE:\s*(' + $numberPattern + ')') { $plusTol = [Math]::Abs((Convert-ToDouble $matches[1])) }
        if ($text -match '(?im)^\s*LOWER TOLERANCE:\s*(' + $numberPattern + ')') { $minusTol = [Math]::Abs((Convert-ToDouble $matches[1])) }
        if ($text -match '(?im)^\s*UPPER SIZE:\s*(' + $numberPattern + ')') { $upperSize = Convert-ToDouble $matches[1]; $hasUpperSize = $true }
        if ($text -match '(?im)^\s*LOWER SIZE:\s*(' + $numberPattern + ')') { $lowerSize = Convert-ToDouble $matches[1]; $hasLowerSize = $true }
        if ($text -match '(?im)^\s*(?:UAME(?: SIZE)?|UNRELATED ACTUAL MATING ENVELOPE(?: SIZE)?):\s*(' + $numberPattern + ')') { $uame = Convert-ToDouble $matches[1]; $hasUame = $true }
        if ($text -match '(?im)^\s*LOCAL SIZE(?:\s*\([^)]*\))?:\s*(' + $numberPattern + ')') { $localSize = Convert-ToDouble $matches[1]; $hasLocalSize = $true }
        if ($text -match '(?im)^\s*FEATURE:\s*(.+?)\s*$') { $feature = $matches[1].Trim() }
        $upperLimit = $nominal + $plusTol
        $lowerLimit = $nominal - $minusTol

        # ASME SIZE commands report UAME and local size as separate evaluated
        # results. Do not substitute the nominal when ISO upper/lower fields are
        # absent; preserve both actual measured values as independent rows.
        if ($hasUame -or $hasLocalSize) {
            if ($hasUame) {
                $deviation = $uame - $nominal
                $outTol = [Math]::Max(0.0, [Math]::Max($uame - $upperLimit, $lowerLimit - $uame))
                $item++
                $rows.Add((New-DimensionRow $item "ASME size $dimensionId - UAME" $dimensionId "UAME" $feature `
                    $nominal $plusTol $minusTol $uame $deviation $outTol 0.0 `
                    $uame $uame 1 ($i + 1) "ProgramTextSize"))
            }
            if ($hasLocalSize) {
                $deviation = $localSize - $nominal
                $outTol = [Math]::Max(0.0, [Math]::Max($localSize - $upperLimit, $lowerLimit - $localSize))
                $item++
                $rows.Add((New-DimensionRow $item "ASME size $dimensionId - local size" $dimensionId "LOCAL_SIZE" $feature `
                    $nominal $plusTol $minusTol $localSize $deviation $outTol 0.0 `
                    $localSize $localSize 1 ($i + 1) "ProgramTextSize"))
            }
            continue
        }

        # ISO SIZE commands provide upper/lower local sizes. Only emit the ISO
        # summary when both evaluated fields exist; otherwise nominal is not a
        # measured result and must not be exported as one.
        if (-not $hasUpperSize -or -not $hasLowerSize) { continue }
        $measured = ($upperSize + $lowerSize) / 2.0
        $deviation = $measured - $nominal
        $outTol = [Math]::Max(0.0, [Math]::Max($upperSize - $upperLimit, $lowerLimit - $lowerSize))
        $item++
        $rows.Add((New-DimensionRow $item "ISO size $dimensionId" $dimensionId "SIZE" $feature `
            $nominal $plusTol $minusTol $measured $deviation $outTol 0.0 `
            $lowerSize $upperSize 1 ($i + 1) "ProgramTextSize"))
    }

    if ($rows.Count -eq 0) {
        throw "No evaluated DIM or SIZE result rows were found in PC-DMIS program: $Path"
    }
    return $rows
}

function Get-PCDMISDimensions {
    $application = $null
    $partProgram = $null
    $commands = $null
    $rows = New-Object System.Collections.Generic.List[object]

    try {
        try {
            $application = Get-PCDMISApplication
        }
        catch {
            throw "No running PC-DMIS Automation instance was found. Run this exporter from the open PC-DMIS program. $($_.Exception.Message)"
        }

        $partProgram = Get-ComProperty $application "ActivePartProgram" $null
        if ($null -eq $partProgram) {
            throw "PC-DMIS has no active part program."
        }

        $commands = Get-ComProperty $partProgram "Commands" $null
        if ($null -eq $commands) {
            throw "The active PC-DMIS part program did not expose its Commands collection."
        }

        $commandCount = [int](Get-ComProperty $commands "Count" 0)
        if ($commandCount -lt 1) {
            throw "The active PC-DMIS part program contains no commands."
        }

        $currentDimensionId = ""
        $currentReportLabel = ""
        $item = 0

        for ($index = 1; $index -le $commandCount; $index++) {
            $command = $null
            $executionCommand = $null
            $dimension = $null
            $commentCommand = $null

            try {
                $command = $commands.Item($index)
                if ($null -eq $command) { continue }

                $isComment = [bool](Get-ComProperty $command "IsComment" $false)
                if ($isComment) {
                    try { $commentCommand = $command.CommentCommand } catch { $commentCommand = $null }
                    if ($null -ne $commentCommand) {
                        $commentType = [int](Get-ComProperty $commentCommand "CommentType" -1)
                        if ($commentType -eq 1) {
                            $commentText = [string](Get-ComProperty $commentCommand "Comment" "")
                            $commentText = ($commentText -replace "`r", " " -replace "`n", " ").Trim()
                            if ($commentText -ne "") { $currentReportLabel = $commentText }
                        }
                    }
                    continue
                }

                $isDimension = [bool](Get-ComProperty $command "IsDimension" $false)
                if (-not $isDimension) { continue }

                $executionCount = [int](Get-ComProperty $command "Count" 1)
                if ($executionCount -lt 1) { $executionCount = 1 }

                # Inside a PC-DMIS LOOP each command has one execution instance per
                # completed iteration. The final instance is the part being exported.
                try { $executionCommand = $command.Item($executionCount) }
                catch { $executionCommand = $command }
                if ($null -eq $executionCommand) { $executionCommand = $command }

                try { $dimension = $executionCommand.DimensionCommand }
                catch { $dimension = $null }
                if ($null -eq $dimension) { continue }

                $ownId = ([string](Get-ComProperty $dimension "ID" "")).Trim()
                $axis = ([string](Get-ComProperty $dimension "AxisLetter" "")).Trim()

                # Location/true-position dimensions have a start command carrying the
                # ID followed by one or more result-axis commands carrying the values.
                if ($ownId -ne "") { $currentDimensionId = $ownId }
                if ($axis -eq "") { continue }

                $dimensionId = $ownId
                if ($dimensionId -eq "") { $dimensionId = $currentDimensionId }
                if ($dimensionId -eq "") { $dimensionId = "DIMENSION_AT_COMMAND_$index" }

                $featureIds = New-Object System.Collections.Generic.List[string]
                foreach ($featureProperty in @("Feat1", "Feat2", "Feat3")) {
                    $featureId = ([string](Get-ComProperty $dimension $featureProperty "")).Trim()
                    if ($featureId -ne "" -and -not $featureIds.Contains($featureId)) {
                        $featureIds.Add($featureId)
                    }
                }

                $item++
                $rows.Add((New-DimensionRow `
                    -Item $item `
                    -ReportLabel $currentReportLabel `
                    -DimensionId $dimensionId `
                    -Axis $axis `
                    -Features ($featureIds -join "; ") `
                    -Nominal (Convert-ToDouble (Get-ComProperty $dimension "Nominal" 0.0)) `
                    -PlusTolerance (Convert-ToDouble (Get-ComProperty $dimension "Plus" 0.0)) `
                    -MinusTolerance (Convert-ToDouble (Get-ComProperty $dimension "Minus" 0.0)) `
                    -Measured (Convert-ToDouble (Get-ComProperty $dimension "Measured" 0.0)) `
                    -Deviation (Convert-ToDouble (Get-ComProperty $dimension "Deviation" 0.0)) `
                    -OutTolerance (Convert-ToDouble (Get-ComProperty $dimension "OutTol" 0.0)) `
                    -Bonus (Convert-ToDouble (Get-ComProperty $dimension "Bonus" 0.0)) `
                    -Minimum (Convert-ToDouble (Get-ComProperty $dimension "Min" 0.0)) `
                    -Maximum (Convert-ToDouble (Get-ComProperty $dimension "Max" 0.0)) `
                    -Execution $executionCount `
                    -CommandIndex $index `
                    -CommandType ([string](Get-ComProperty $executionCommand "Type" ""))
                ))
            }
            finally {
                Release-ComObject $commentCommand
                Release-ComObject $dimension
                if ($executionCommand -ne $command) { Release-ComObject $executionCommand }
                Release-ComObject $command
            }
        }

        if ($rows.Count -eq 0) {
            throw "No evaluated PC-DMIS dimension result commands were found. Ensure the exporter is called after the dimensions have executed."
        }

        return $rows
    }
    finally {
        Release-ComObject $commands
        Release-ComObject $partProgram
        Release-ComObject $application
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Write-CsvFallback {
    param(
        [Parameter(Mandatory = $true)]$Rows,
        [Parameter(Mandatory = $true)][string]$Path
    )

    $Rows | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
}

function Write-ExcelReport {
    param(
        [Parameter(Mandatory = $true)]$Rows,
        [Parameter(Mandatory = $true)][string]$Path,
        [string]$OperatorName,
        [string]$Job,
        [string]$Serial,
        [string]$Part,
        [string]$Rev
    )

    $excel = $null
    $workbook = $null
    $sheet = $null
    $range = $null
    $stage = "starting Excel"

    try {
        $stage = "creating Excel application"
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $stage = "creating workbook"
        $workbook = $excel.Workbooks.Add()
        $sheet = $workbook.Worksheets.Item(1)
        $sheet.Name = "Inspection Report"

        $headers = @(
            "Item", "Drawing / Characteristic", "Dimension", "Axis", "Features",
            "Nominal", "+Tol", "-Tol", "Measured", "Deviation", "Out Tol",
            "Bonus", "Minimum", "Maximum", "Status", "Execution",
            "Command Index", "Command Type"
        )

        $lastColumn = $headers.Count
        $headerRow = 9
        $dataStartRow = $headerRow + 1

        $stage = "writing title and metadata"
        $sheet.Cells.Item(1, 1).Value2 = "PC-DMIS Inspection Report - Fallback Export"
        $titleRange = $sheet.Range($sheet.Cells.Item(1, 1), $sheet.Cells.Item(1, $lastColumn))
        $titleRange.Merge()
        $titleRange.Font.Bold = $true
        $titleRange.Font.Size = 16
        $titleRange.Font.Color = 0xFFFFFF
        $titleRange.Interior.Color = 0x6B3A00
        $titleRange.HorizontalAlignment = -4108
        $titleRange.RowHeight = 26

        $sheet.Cells.Item(3, 1).Value2 = "Part Number"
        $sheet.Cells.Item(3, 2).Value2 = $Part
        $sheet.Cells.Item(3, 4).Value2 = "Revision"
        $sheet.Cells.Item(3, 5).Value2 = $Rev
        $sheet.Cells.Item(4, 1).Value2 = "Job Number"
        $sheet.Cells.Item(4, 2).Value2 = $Job
        $sheet.Cells.Item(4, 4).Value2 = "Serial Number"
        $sheet.Cells.Item(4, 5).Value2 = $Serial
        $sheet.Cells.Item(5, 1).Value2 = "Operator"
        $sheet.Cells.Item(5, 2).Value2 = $OperatorName
        $sheet.Cells.Item(5, 4).Value2 = "Exported"
        $sheet.Cells.Item(5, 5).Value2 = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        $sheet.Cells.Item(6, 1).Value2 = "Result Count"
        $sheet.Cells.Item(6, 2).Value2 = [int]$Rows.Count
        $sheet.Cells.Item(6, 4).Value2 = "Failed"
        $sheet.Cells.Item(6, 5).Value2 = [int](@($Rows | Where-Object { $_.Status -eq "FAIL" }).Count)

        $metaLabels = $sheet.Range("A3:A6,D3:D6")
        $metaLabels.Font.Bold = $true
        $metaLabels.Interior.Color = 0xD9EAF7

        $stage = "writing column headers"
        for ($column = 1; $column -le $lastColumn; $column++) {
            $sheet.Cells.Item($headerRow, $column).Value2 = $headers[$column - 1]
        }

        $headerRange = $sheet.Range($sheet.Cells.Item($headerRow, 1), $sheet.Cells.Item($headerRow, $lastColumn))
        $headerRange.Font.Bold = $true
        $headerRange.Font.Color = 0xFFFFFF
        $headerRange.Interior.Color = 0x6B3A00
        $headerRange.HorizontalAlignment = -4108
        $headerRange.VerticalAlignment = -4108
        $headerRange.WrapText = $true

        $stage = "writing dimension rows"
        $rowNumber = $dataStartRow
        foreach ($dimensionRow in $Rows) {
            $values = @(
                [double]$dimensionRow.Item,
                $dimensionRow.ReportLabel,
                $dimensionRow.Dimension,
                $dimensionRow.Axis,
                $dimensionRow.Features,
                $dimensionRow.Nominal,
                $dimensionRow.PlusTolerance,
                $dimensionRow.MinusTolerance,
                $dimensionRow.Measured,
                $dimensionRow.Deviation,
                $dimensionRow.OutTolerance,
                $dimensionRow.Bonus,
                $dimensionRow.Minimum,
                $dimensionRow.Maximum,
                $dimensionRow.Status,
                [double]$dimensionRow.Execution,
                [double]$dimensionRow.CommandIndex,
                $dimensionRow.CommandType
            )

            for ($column = 1; $column -le $lastColumn; $column++) {
                try {
                    $cellValue = $values[$column - 1]
                    # Passing mixed numeric values through a PowerShell object array
                    # can make the Excel COM binder select its String overload. An
                    # invariant string is safe here: Excel converts numeric strings
                    # back into numeric cells when Value2 is assigned.
                    if ($null -eq $cellValue) {
                        $excelValue = ""
                    }
                    elseif ($cellValue -is [IFormattable] -and -not ($cellValue -is [string])) {
                        $excelValue = $cellValue.ToString($null, [Globalization.CultureInfo]::InvariantCulture)
                    }
                    else {
                        $excelValue = [string]$cellValue
                    }
                    $sheet.Cells.Item($rowNumber, $column).Value2 = $excelValue
                }
                catch {
                    $valueType = "null"
                    if ($null -ne $cellValue) { $valueType = $cellValue.GetType().FullName }
                    throw "Could not write row $rowNumber column $column ($($headers[$column - 1])) value type $valueType. $($_.Exception.Message)"
                }
            }

            $statusCell = $sheet.Cells.Item($rowNumber, 15)
            $statusCell.Font.Bold = $true
            $statusCell.HorizontalAlignment = -4108
            if ($dimensionRow.Status -eq "FAIL") {
                $statusCell.Interior.Color = 0xC7C6FF
                $statusCell.Font.Color = 0x0000C0
            }
            else {
                $statusCell.Interior.Color = 0xCEEFCE
                $statusCell.Font.Color = 0x006000
            }

            $rowNumber++
        }

        $stage = "formatting table"
        $lastDataRow = $rowNumber - 1
        $dataRange = $sheet.Range($sheet.Cells.Item($headerRow, 1), $sheet.Cells.Item($lastDataRow, $lastColumn))
        $dataRange.Borders.LineStyle = 1
        $dataRange.Borders.Weight = 2
        $dataRange.VerticalAlignment = -4108

        $numericRange = $sheet.Range($sheet.Cells.Item($dataStartRow, 6), $sheet.Cells.Item($lastDataRow, 14))
        $numericRange.NumberFormat = "0.000"

        $sheet.Range($sheet.Cells.Item($dataStartRow, 1), $sheet.Cells.Item($lastDataRow, 1)).HorizontalAlignment = -4108
        $sheet.Range($sheet.Cells.Item($dataStartRow, 4), $sheet.Cells.Item($lastDataRow, 4)).HorizontalAlignment = -4108
        $sheet.Range($sheet.Cells.Item($dataStartRow, 15), $sheet.Cells.Item($lastDataRow, 18)).HorizontalAlignment = -4108

        $sheet.Columns.Item(1).ColumnWidth = 7
        $sheet.Columns.Item(2).ColumnWidth = 38
        $sheet.Columns.Item(3).ColumnWidth = 18
        $sheet.Columns.Item(4).ColumnWidth = 8
        $sheet.Columns.Item(5).ColumnWidth = 28
        for ($column = 6; $column -le 14; $column++) { $sheet.Columns.Item($column).ColumnWidth = 12 }
        $sheet.Columns.Item(15).ColumnWidth = 10
        $sheet.Columns.Item(16).ColumnWidth = 10
        $sheet.Columns.Item(17).ColumnWidth = 13
        $sheet.Columns.Item(18).ColumnWidth = 14

        $sheet.Range($sheet.Cells.Item($dataStartRow, 2), $sheet.Cells.Item($lastDataRow, 5)).WrapText = $true
        $headerRange.AutoFilter() | Out-Null

        $stage = "configuring page layout"
        $sheet.Activate() | Out-Null
        $excel.ActiveWindow.SplitRow = $headerRow
        $excel.ActiveWindow.FreezePanes = $true
        $sheet.PageSetup.Orientation = 2
        $sheet.PageSetup.Zoom = $false
        $sheet.PageSetup.FitToPagesWide = 1
        $sheet.PageSetup.FitToPagesTall = $false
        $sheet.PageSetup.PrintTitleRows = ('$' + $headerRow + ':$' + $headerRow)

        # 51 = xlOpenXMLWorkbook (.xlsx). Save to a temporary workbook first,
        # then replace the destination so an interrupted export cannot leave a
        # half-written final report.
        $temporaryPath = [IO.Path]::Combine(
            [IO.Path]::GetDirectoryName($Path),
            ([IO.Path]::GetFileNameWithoutExtension($Path) + ".partial-" + [Guid]::NewGuid().ToString("N") + ".xlsx")
        )

        $stage = "saving temporary workbook"
        $workbook.SaveAs($temporaryPath, 51)
        $workbook.Close($false)
        Release-ComObject $workbook
        $workbook = $null

        if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
        $stage = "moving completed workbook"
        Move-Item -LiteralPath $temporaryPath -Destination $Path -Force
    }
    catch {
        throw "Excel export failed while $stage. $($_.Exception.Message)"
    }
    finally {
        if ($null -ne $workbook) {
            try { $workbook.Close($false) } catch { }
        }
        if ($null -ne $excel) {
            try { $excel.Quit() } catch { }
        }
        Release-ComObject $range
        Release-ComObject $sheet
        Release-ComObject $workbook
        Release-ComObject $excel
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

try {
    $requestPath = Join-Path $PSScriptRoot "export-request.txt"
    if (-not $Mock -and [string]::IsNullOrWhiteSpace($OutputPath) -and (Test-Path -LiteralPath $requestPath)) {
        $request = @(Get-Content -LiteralPath $requestPath)
        if ($request.Count -lt 7) {
            throw "The export request must contain at least seven lines: output, operator, job, serial, part, revision, and program path. An optional eighth line supplies the PC-DMIS BASIC bridge path."
        }
        $OutputPath = $request[0].Trim()
        $Operator = $request[1].Trim()
        $JobNumber = $request[2].Trim()
        $SerialNumber = $request[3].Trim()
        $PartNumber = $request[4].Trim()
        $Revision = $request[5].Trim()
        $ProgramPath = $request[6].Trim()
        if ($request.Count -ge 8) { $BridgePath = $request[7].Trim() }
    }

    if (-not $Mock -and [string]::IsNullOrWhiteSpace($OutputPath)) {
        # Legacy Automation integration retained for installations whose COM
        # registration is healthy. The request-file route does not require COM.
        $OutputPath = Get-PCDMISVariableText "EXPORT_FILE"
        if ([string]::IsNullOrWhiteSpace($Operator)) { $Operator = Get-PCDMISVariableText "EXPORT_OPERATOR" }
        if ([string]::IsNullOrWhiteSpace($JobNumber)) { $JobNumber = Get-PCDMISVariableText "EXPORT_JOB" }
        if ([string]::IsNullOrWhiteSpace($SerialNumber)) { $SerialNumber = Get-PCDMISVariableText "EXPORT_SERIAL" }
        if ([string]::IsNullOrWhiteSpace($PartNumber)) { $PartNumber = Get-PCDMISVariableText "EXPORT_PART" }
        if ([string]::IsNullOrWhiteSpace($Revision)) { $Revision = Get-PCDMISVariableText "EXPORT_REV" }
        if ([string]::IsNullOrWhiteSpace($ProgramPath)) { $ProgramPath = Get-PCDMISVariableText "EXPORT_PROGRAM" }
    }

    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        throw "No output path was supplied. Pass an .xlsx path or define the PC-DMIS variable EXPORT_FILE before calling the exporter."
    }

    if ([IO.Path]::GetExtension($OutputPath) -ne ".xlsx") {
        $OutputPath = [IO.Path]::ChangeExtension($OutputPath, ".xlsx")
    }
    $OutputPath = [IO.Path]::GetFullPath($OutputPath)
    $outputDirectory = [IO.Path]::GetDirectoryName($OutputPath)
    if ([string]::IsNullOrWhiteSpace($outputDirectory)) {
        throw "The report output path has no directory."
    }
    if (-not (Test-Path -LiteralPath $outputDirectory)) {
        New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
    }

    $logPath = [IO.Path]::ChangeExtension($OutputPath, ".export.log")
    $csvPath = [IO.Path]::ChangeExtension($OutputPath, ".csv")

    if ($Mock) {
        $dimensionRows = @(Get-MockDimensions)
    }
    elseif (-not [string]::IsNullOrWhiteSpace($BridgePath)) {
        $BridgePath = [IO.Path]::GetFullPath($BridgePath)
        $dimensionRows = @(Get-PCDMISBridgeDimensions -Path $BridgePath)
        if (-not [string]::IsNullOrWhiteSpace($ProgramPath)) {
            $ProgramPath = [IO.Path]::GetFullPath($ProgramPath)
            $programRows = @(Get-PCDMISProgramDimensions -Path $ProgramPath)
            $sizeRows = @($programRows | Where-Object { $_.CommandType -eq 'ProgramTextSize' })
            foreach ($sizeRow in $sizeRows) {
                $sizeRow.Item = $dimensionRows.Count + 1
                $dimensionRows += $sizeRow
            }
        }
    }
    elseif (-not [string]::IsNullOrWhiteSpace($ProgramPath)) {
        $ProgramPath = [IO.Path]::GetFullPath($ProgramPath)
        $dimensionRows = @(Get-PCDMISProgramDimensions -Path $ProgramPath)
    }
    else {
        $dimensionRows = @(Get-PCDMISDimensions)
    }

    # CSV is deliberately written before Excel. It remains a readable emergency
    # result if Excel COM fails on a damaged customer installation.
    Write-CsvFallback -Rows $dimensionRows -Path $csvPath
    Write-ExcelReport -Rows $dimensionRows -Path $OutputPath `
        -OperatorName $Operator -Job $JobNumber -Serial $SerialNumber `
        -Part $PartNumber -Rev $Revision

    @(
        "SUCCESS"
        "Time=$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))"
        "Workbook=$OutputPath"
        "CSV=$csvPath"
        "SourceProgram=$ProgramPath"
        "Bridge=$BridgePath"
        "Rows=$($dimensionRows.Count)"
        "Failed=$(@($dimensionRows | Where-Object { $_.Status -eq 'FAIL' }).Count)"
    ) | Set-Content -LiteralPath $logPath -Encoding UTF8

    Write-Output "PC-DMIS fallback report saved: $OutputPath"
    exit 0
}
catch {
    $message = $_.Exception.Message
    try {
        $fixedFailureLog = Join-Path $PSScriptRoot "last-export-error.log"
        @(
            "FAILED"
            "Time=$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))"
            "Message=$message"
            "Details=$($_ | Out-String)"
        ) | Set-Content -LiteralPath $fixedFailureLog -Encoding UTF8

        if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
            $failureDirectory = [IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($OutputPath))
            if (-not [string]::IsNullOrWhiteSpace($failureDirectory) -and -not (Test-Path -LiteralPath $failureDirectory)) {
                New-Item -ItemType Directory -Path $failureDirectory -Force | Out-Null
            }
            $failureLog = [IO.Path]::ChangeExtension([IO.Path]::GetFullPath($OutputPath), ".export-error.log")
            @(
                "FAILED"
                "Time=$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))"
                "Message=$message"
                "Details=$($_ | Out-String)"
            ) | Set-Content -LiteralPath $failureLog -Encoding UTF8
        }
    }
    catch { }

    Write-Error $message
    exit 1
}
