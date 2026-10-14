# Abre un DOCX/DOC con Microsoft Word (invisible, solo lectura) y lo exporta a PDF.
# Solo para QA local; la app no depende de Word.
# Uso: powershell -File scripts/word-a-pdf.ps1 entrada.docx salida.pdf
param([string]$Entrada, [string]$Salida)
$entradaAbs = (Resolve-Path $Entrada).Path
$salidaAbs = [System.IO.Path]::GetFullPath($Salida)
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
  $doc = $word.Documents.Open($entradaAbs, $false, $true, $false)
  Write-Output ("Paginas segun Word: " + $doc.ComputeStatistics(2))
  $doc.ExportAsFixedFormat($salidaAbs, 17)
  $doc.Close(0)
} finally {
  $word.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
Write-Output "OK $salidaAbs"
