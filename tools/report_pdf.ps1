param([string]$Path, [string]$Pdf, [int]$SheetIndex = 12)
$x = New-Object -ComObject Excel.Application
$x.Visible = $false
$x.DisplayAlerts = $false
try {
    $wb = $x.Workbooks.Open($Path, 0, $true)
    $x.CalculateFull()
    $ws = $wb.Worksheets.Item($SheetIndex)
    "sheet: " + $ws.Index + " rows used: " + $ws.UsedRange.Rows.Count
    $ws.Unprotect()
    foreach ($r in (16..55) + (59..98)) { if ($ws.Cells.Item($r, 1).Text -eq '') { $ws.Rows.Item($r).Hidden = $true } }
    $ws.ExportAsFixedFormat(0, $Pdf)
    $wb.Close($false)
    "pdf-ok"
} finally {
    $x.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($x)
}
