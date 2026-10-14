# Guarda un DOCX como .doc (97-2003), .rtf y .odt usando Microsoft Word (solo para generar archivos de prueba).
param([string]$Entrada)
$abs = (Resolve-Path $Entrada).Path
$base = [System.IO.Path]::ChangeExtension($abs, $null).TrimEnd('.')
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
  $doc = $word.Documents.Open($abs, $false, $true, $false)
  $doc.SaveAs2("$base.doc", 0)
  $doc.SaveAs2("$base.rtf", 6)
  $doc.SaveAs2("$base.odt", 23)
  $doc.Close(0)
} finally {
  $word.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
Write-Output "OK $base.{doc,rtf,odt}"
