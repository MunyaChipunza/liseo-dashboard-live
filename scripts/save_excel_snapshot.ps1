[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SourcePath,

    [Parameter(Mandatory = $true)]
    [string]$TargetPath
)

$ErrorActionPreference = "Stop"

$resolvedSource = [System.IO.Path]::GetFullPath($SourcePath)
$resolvedTarget = [System.IO.Path]::GetFullPath($TargetPath)

$targetDir = Split-Path -Parent $resolvedTarget
if ($targetDir -and -not (Test-Path -LiteralPath $targetDir)) {
    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
}

function Save-CopyFromLiveWorkbook {
    param(
        [Parameter(Mandatory = $true)]
        [string]$WorkbookPath,

        [Parameter(Mandatory = $true)]
        [string]$TargetPath
    )

    $excelApp = $null
    $candidate = $null

    try {
        try {
            $excelApp = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        }
        catch {
            return $false
        }

        $workbooks = $excelApp.Workbooks
        for ($index = 1; $index -le $workbooks.Count; $index++) {
            $candidate = $workbooks.Item($index)
            $candidatePath = $candidate.FullName
            if (-not $candidatePath) {
                if ($candidate -ne $null) {
                    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($candidate)
                    $candidate = $null
                }
                continue
            }

            $samePath = [string]::Equals(
                [System.IO.Path]::GetFullPath($candidatePath),
                $WorkbookPath,
                [System.StringComparison]::OrdinalIgnoreCase
            )

            if (-not $samePath) {
                if ($candidate -ne $null) {
                    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($candidate)
                    $candidate = $null
                }
                continue
            }

            for ($attempt = 0; $attempt -lt 4; $attempt++) {
                try {
                    $candidate.SaveCopyAs($TargetPath)
                    return $true
                }
                catch {
                    if ($attempt -eq 3 -or $_.Exception.Message -notmatch '0x800AC472') {
                        throw
                    }
                    Start-Sleep -Milliseconds 750
                }
            }
        }

        return $false
    }
    finally {
        if ($candidate -ne $null) {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($candidate)
        }
        if ($excelApp -ne $null) {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excelApp)
        }
    }
}

if (Save-CopyFromLiveWorkbook -WorkbookPath $resolvedSource -TargetPath $resolvedTarget) {
    return
}

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.DisplayAlerts = $false
    $excel.Visible = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false

    Start-Sleep -Milliseconds 250
    $workbook = $excel.Workbooks.Open($resolvedSource)

    $workbook.SaveCopyAs($resolvedTarget)
}
finally {
    if ($workbook -ne $null) {
        $workbook.Close($false) | Out-Null
    }

    if ($workbook -ne $null) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }

    if ($excel -ne $null) {
        $excel.Quit() | Out-Null
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
