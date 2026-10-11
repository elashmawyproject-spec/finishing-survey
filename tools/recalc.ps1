param([string]$Path)
$x = New-Object -ComObject Excel.Application
$x.Visible = $false
$x.DisplayAlerts = $false
try {
    $wb = $x.Workbooks.Open($Path)
    $x.CalculateFull()
    "sheets: " + $wb.Worksheets.Count
    $wb.Save()
    $wb.Close($false)
    "recalc-ok"
} finally {
    $x.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($x)
}
