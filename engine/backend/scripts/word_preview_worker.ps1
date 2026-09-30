# Persistent, single-threaded Word worker.
# It never edits DOCX structure. The caller performs any Python/XML pagination
# changes on disk between requests. Each request opens one document, performs
# a field pass or a screen preview export, then closes it. One JSON request per
# stdin line; one JSON response per stdout line.
# ASCII only: Windows PowerShell 5.1 treats BOM-less scripts as ANSI.

$word = $null
try {
    [Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false)
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = [Microsoft.Office.Interop.Word.WdAlertLevel]::wdAlertsNone
    [Console]::Out.WriteLine('{"ready":true}')

    while (($line = [Console]::In.ReadLine()) -ne $null) {
        $doc = $null
        try {
            $request = $line | ConvertFrom-Json
            $doc = $word.Documents.Open([string]$request.input, $false, $false)
            if ($request.op -eq 'preview') {
                $doc.Repaginate()
                $doc.Range().Fields.Update() | Out-Null
                $doc.Repaginate()
                foreach ($toc in $doc.TablesOfContents) {
                    $toc.UpdatePageNumbers() | Out-Null
                }
                # A preview can optionally retain its field-result cache. The
                # API stores this only in the separate fast-cache copy; final
                # pagination never mutates the immutable source cache.
                if ($request.persist_fields) { $doc.Save() }
                # 1 = wdExportOptimizeForOnScreen. This branch is preview-only.
                $doc.ExportAsFixedFormat([string]$request.output, 17, $false, 1)
                $doc.Close(0)
                $doc = $null
                [Console]::Out.WriteLine('{"ok":true}')
                continue
            }

            if ($request.op -eq 'export_full') {
                # 17 = wdFormatPDF. The final DOCX has already had its fields
                # updated and pagination saved by a preceding pass.
                $doc.SaveAs2([string]$request.output, 17)
                $doc.Close(0)
                $doc = $null
                [Console]::Out.WriteLine('{"ok":true}')
                continue
            }

            if ($request.op -ne 'pass') { throw "unsupported worker operation: $($request.op)" }
            if ($request.update_fields) {
                $doc.Range().Fields.Update() | Out-Null
                $doc.Repaginate()
            }
            if ($request.update_page_numbers) {
                $doc.Repaginate()
                foreach ($toc in $doc.TablesOfContents) {
                    $toc.UpdatePageNumbers() | Out-Null
                }
            }
            if ($request.update_fields -or $request.update_page_numbers) { $doc.Save() }
            $sections = @{}
            $total = $null
            if ($request.measure) {
                $doc.Repaginate()
                for ($i = 1; $i -le $doc.Sections.Count; $i++) {
                    $rng = $doc.Sections.Item($i).Range
                    $rng.Collapse(1)
                    $sections["$i"] = $rng.Information(3)
                }
                $total = $doc.ComputeStatistics(2)
            }
            $doc.Close(0)
            $doc = $null
            [Console]::Out.WriteLine((@{ok=$true;sections=$sections;pages=$total} | ConvertTo-Json -Compress))
        } catch {
            if ($doc -ne $null) { try { $doc.Close(0) } catch {} }
            $escaped = ($_ | Out-String).Trim().Replace('"', '\"').Replace("`r", ' ').Replace("`n", ' ')
            [Console]::Out.WriteLine('{"ok":false,"error":"' + $escaped + '"}')
        }
    }
} finally {
    if ($word -ne $null) {
        try { $word.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
}
