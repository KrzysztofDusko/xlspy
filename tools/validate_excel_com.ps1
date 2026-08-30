[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0, ValueFromRemainingArguments = $true)]
    [string[]] $Path,

    [switch] $RefreshPivots
)

$ErrorActionPreference = "Stop"
$failures = [System.Collections.Generic.List[string]]::new()
$excel = $null

function Release-ComObject {
    param([object] $Value)
    if ($null -ne $Value -and [System.Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false
    $workbooks = $excel.Workbooks

    foreach ($candidate in $Path) {
        $workbook = $null
        $worksheets = $null
        try {
            $fullPath = [System.IO.Path]::GetFullPath($candidate)
            if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) {
                throw "File does not exist: $fullPath"
            }

            # Open read-only with Excel's normal loader.  A corrupt-load mode
            # is deliberately not used: silent repair would make validation
            # report a damaged file as valid.
            $workbook = $workbooks.Open($fullPath, 0, $true)
            $worksheets = $workbook.Worksheets
            $sheetCount = [int]$worksheets.Count
            if ($sheetCount -lt 1) {
                throw "Workbook has no worksheets: $fullPath"
            }

            for ($sheetIndex = 1; $sheetIndex -le $sheetCount; $sheetIndex++) {
                $worksheet = $null
                $usedRange = $null
                try {
                    $worksheet = $worksheets.Item($sheetIndex)
                    $usedRange = $worksheet.UsedRange
                    # Force Excel to materialise the complete used region.
                    $null = $usedRange.Value2
                    if ($RefreshPivots) {
                        $pivotTables = $worksheet.PivotTables()
                        for ($pivotIndex = 1; $pivotIndex -le [int]$pivotTables.Count; $pivotIndex++) {
                            $pivotTable = $null
                            try {
                                $pivotTable = $pivotTables.Item($pivotIndex)
                                $null = $pivotTable.RefreshTable()
                            }
                            finally {
                                Release-ComObject $pivotTable
                            }
                        }
                        Release-ComObject $pivotTables
                    }
                }
                finally {
                    Release-ComObject $usedRange
                    Release-ComObject $worksheet
                }
            }

            Write-Output ("OK`t{0}`tworksheets={1}" -f $fullPath, $sheetCount)
        }
        catch {
            $failures.Add(("FAIL`t{0}`t{1}" -f $candidate, $_.Exception.Message))
        }
        finally {
            if ($null -ne $workbook) {
                try { $workbook.Close($false) } catch { }
                Release-ComObject $workbook
            }
            Release-ComObject $worksheets
            $worksheets = $null
        }
    }
}
catch {
    $failures.Add(("Excel COM unavailable: {0}" -f $_.Exception.Message))
}
finally {
    if ($null -ne $excel) {
        try { $excel.Quit() } catch { }
        Release-ComObject $workbooks
        Release-ComObject $excel
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

foreach ($failure in $failures) {
    Write-Error $failure
}
if ($failures.Count -gt 0) {
    exit 1
}
exit 0
