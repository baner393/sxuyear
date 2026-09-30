# One Word session doing selected passes over a docx, in this order:
#   -UpdateFields      full field update (rebuilds the TOC field), then save
#   -UpdatePageNumbers refresh TOC page numbers only, then save
#   -Measure           print "N:page" per section + "PAGES:total" (physical pages)
# Replaces update_fields.ps1 + measure_sections.ps1: merging passes into one
# session saves a full Word cold start (~8s) per merged call.
# NOTE: keep this file ASCII-only (PowerShell 5.1 reads BOM-less files as ANSI).
param(
    [Parameter(Mandatory=$true)][string]$Docx,
    [switch]$UpdateFields,
    [switch]$UpdatePageNumbers,
    [switch]$Measure
)

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = [Microsoft.Office.Interop.Word.WdAlertLevel]::wdAlertsNone

    $readOnly = -not ($UpdateFields -or $UpdatePageNumbers)
    $doc = $word.Documents.Open($Docx, $false, $readOnly)

    if ($UpdateFields) {
        # Range().Fields.Update() already rebuilds the TOC field (it lives in
        # the main story); the old script updated TOCs a second time for
        # nothing (verified equivalent by A/B on real output).
        $doc.Range().Fields.Update() | Out-Null
        $doc.Repaginate()
    }
    if ($UpdatePageNumbers) {
        $doc.Repaginate()
        foreach ($toc in $doc.TablesOfContents) {
            $toc.UpdatePageNumbers() | Out-Null
        }
    }
    if ($UpdateFields -or $UpdatePageNumbers) {
        $doc.Save()
    }
    if ($Measure) {
        $doc.Repaginate()
        # wdActiveEndPageNumber = 3: physical page counted from document start
        # (NOT the displayed number, which follows per-section numbering restarts)
        for ($i = 1; $i -le $doc.Sections.Count; $i++) {
            $rng = $doc.Sections.Item($i).Range
            $rng.Collapse(1)
            Write-Output ("{0}:{1}" -f $i, $rng.Information(3))
        }
        # wdStatisticPages = 2: real total page count for the UI
        Write-Output ("PAGES:{0}" -f $doc.ComputeStatistics(2))
    }

    $doc.Close(0)
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    Write-Output "OK"
} catch {
    Write-Error "ERROR (line $($_.InvocationInfo.ScriptLineNumber)): $_"
    exit 1
}
