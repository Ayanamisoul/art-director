param([string]$in, [string]$out)
$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open($in, $true, $false, $false)
$pres.SaveAs($out, 32)  # 32 = ppSaveAsPDF
$pres.Close()
$ppt.Quit()
Write-Output "OK $out"
