# Word PowerShell: Convert DOCX to PDF.
#   -UpdateFields  rebuild fields/TOC in-session before export (never saved back).
#                  Pass it for docx that skipped the pagination pass (fast
#                  preview); omit for freshly converted docx whose fields were
#                  already updated and saved by word_pass.ps1 (saves 3-8s).
#   -Preview       screen-optimized PDF for the interactive preview only.
#                  The deliverable PDF keeps SaveAs2's original-quality output.
# NOTE: keep this file ASCII-only. PowerShell 5.1 reads BOM-less scripts as ANSI,
# so non-ASCII comments can corrupt adjacent lines.
param(
    [Parameter(Mandatory=$true)][string]$InputDocx,
    [Parameter(Mandatory=$true)][string]$OutputPdf,
    [switch]$UpdateFields,
    [switch]$Preview
)

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = [Microsoft.Office.Interop.Word.WdAlertLevel]::wdAlertsNone

    # Open writable (changes are never saved back) so TOC/PAGE fields can update
    $doc = $word.Documents.Open($InputDocx, $false, $false)

    if ($UpdateFields) {
        $doc.Repaginate()
        $doc.Range().Fields.Update() | Out-Null
        $doc.Repaginate()
        foreach ($toc in $doc.TablesOfContents) {
            $toc.UpdatePageNumbers() | Out-Null
        }
    }

    if ($Preview) {
        # 1 = wdExportOptimizeForOnScreen. Measured against the real 36-page
        # preview: 37% smaller, with no text/layout change at preview scale.
        $doc.ExportAsFixedFormat($OutputPdf, 17, $false, 1)
    } else {
        $doc.SaveAs2($OutputPdf, 17)  # 17 = wdFormatPDF (deliverable)
    }
    $doc.Close(0)  # 0 = wdDoNotSaveChanges
    $word.Quit()

    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    Write-Output "OK"
} catch {
    Write-Error "ERROR (line $($_.InvocationInfo.ScriptLineNumber)): $_"
    exit 1
}
