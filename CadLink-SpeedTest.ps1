# Uncomment this if scriptblock to force running as administrator and bypass local execution policy
#if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) 
#{ 
#    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
#    exit 
#}

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13

Clear-Host

[string]$CadLinkAssembly = 'C:\Program Files\Cadwell\CadLinkClientService\CadLink.Common.dll'
Add-Type -Path $CadLinkAssembly

$pathToXml = "C:\ProgramData\Cadwell\CadLinkClientSettings.xml"
$ClientSettingsXml = New-Object XML
$ClientSettingsXml.Load($($pathToXml))
$StreamingBatchSizeBytes = $ClientSettingsXml.SelectSingleNode("//StreamingBatchSizeBytes")
[int]$batchSize = $StreamingBatchSizeBytes.InnerText

$date = Get-Date -Format "yyyy/MM/dd HH:mm:ss"
$logresults = 10
$timer = 0
$Latency = 0
$totalLatency = @()
$sumTotalLatency = 0
$UploadSpeed = 0
$totalUpload = @()
$sumTotalUpload = 0
$DownloadSpeed = 0
$totalDownload = @()
$sumTotalDownload = 0

Write-Host "Test started at $(Get-Date -Format `"HH:mm:ss yyyy/MM/dd`")"
Write-Host "The time needed to run this test will vary based on latency."
Write-Host "If latency is high, this test will take a long time to complete."
Write-Host ""

do {
    
    Start-Sleep 1
    
    $logresults -= 1
    $timer += 1 
    
    $runtest = [CadLink.Common.Communication.NetworkDoctor]::RunTestsAndLogResults()  
    
    $Latency = $runtest | Select-Object -ExpandProperty Latency
    $UploadSpeed = $runtest | Select-Object -ExpandProperty UploadSpeed
    $DownloadSpeed = $runtest | Select-Object -ExpandProperty DownloadSpeed    

    $totalLatency += $Latency
    $totalUpload += $UploadSpeed
    $totalDownload += $DownloadSpeed

    $sumTotalLatency = ($totalLatency | Measure-Object -Sum).Sum
    $sumTotalUpload = ($totalUpload | Measure-Object -Sum).Sum
    $sumTotalDownload = ($totalDownload | Measure-Object -Sum).Sum
    
    $averageLatency = [math]::Round(($sumTotalLatency / $totalLatency.Count),2)
    $avgUp = [math]::Round(($sumTotalUpload / $totalUpload.Count),2)
    $avgDown = [math]::Round(($sumTotalDownload / $totalDownload.Count),2)
} 

while (

    $logresults -ge 1

)

Write-Host "Test completed at $(Get-Date -Format `"HH:mm:ss yyyy/MM/dd`")"
Write-Host "Estimated Test Duration: $([math]::Round($($timer/60),2)) minute(s)"
Write-Host ""
Write-Host "Average Latency:" -BackgroundColor DarkGray -ForegroundColor Yellow
if (($averageLatency -gt 0) -and ($averageLatency -le 20)) {
    
    Write-Host "  Low: $averageLatency ms" #-ForegroundColor Green

} elseif (($averageLatency -gt 20) -and ($averageLatency -le 50)) {
    
    Write-Host "  Medium: $averageLatency ms" #-ForegroundColor Yellow

} elseif (($averageLatency -gt 50) -and ($averageLatency -le 75)) {
    
    Write-Host "  High: $averageLatency ms" #-ForegroundColor Yellow

} elseif ($averageLatency -gt 75) {

    Write-Host "  Very High: $averageLatency ms" #-ForegroundColor Red    

}
Write-Host ""
Write-Host "Average Upload Speed:" -BackgroundColor DarkGray -ForegroundColor Yellow
Write-Host "  (Megabits Per Second):  $avgUp"
Write-Host "  (Megabytes Per Seconds): $($avgUp/8)"
Write-Host ""
Write-Host "Average Download Speed:" -BackgroundColor DarkGray -ForegroundColor Yellow
Write-Host "  (Megabits Per Second):  $avgDown"
Write-Host "  (Megabytes Per Seconds): $($avgDown/8)"
Write-Host ""
Write-Host "Max data upload capacity:" -BackgroundColor DarkGray -ForegroundColor Yellow
Write-Host "  (Megabits Per Minute):  $($avgUp*60)"
Write-Host "  (Megabytes Per Minute): $(($avgUp/8)*60)"
Write-Host ""
Write-Host "Estimated upload times for a given file size:" -BackgroundColor DarkGray -ForegroundColor Yellow
Write-Host "  This client's sync batch size $(($batchSize / 1024 / 1024)) MB: $([math]::Round(((50 * 8) / $avgUp),2)) seconds  =  $([math]::Round((((50 * 8) / $avgUp)/60),2)) minutes  =  $([math]::Round(((((50 * 8) / $avgUp)/60)/60),2)) hours" -ForegroundColor Green
Write-Host "  50 MB:  $([math]::Round(((50 * 8) / $avgUp),2)) seconds  =  $([math]::Round((((50 * 8) / $avgUp)/60),2)) minutes  =  $([math]::Round(((((50 * 8) / $avgUp)/60)/60),2)) hours"
Write-Host "  100 MB:  $([math]::Round(((100 * 8) / $avgUp),2)) seconds  =  $([math]::Round((((100 * 8) / $avgUp)/60),2)) minutes  =  $([math]::Round(((((100 * 8) / $avgUp)/60)/60),2)) hours"
Write-Host "  250 MB:  $([math]::Round(((250 * 8) / $avgUp),2)) seconds  =  $([math]::Round((((250 * 8) / $avgUp)/60),2)) minutes  =  $([math]::Round(((((250 * 8) / $avgUp)/60)/60),2)) hours"
Write-Host "  500 MB:  $([math]::Round(((500 * 8) / $avgUp),2)) seconds  =  $([math]::Round((((500 * 8) / $avgUp)/60),2)) minutes  =  $([math]::Round(((((500 * 8) / $avgUp)/60)/60),2)) hours"
Write-Host "  750 MB:  $([math]::Round(((750 * 8) / $avgUp),2)) seconds  =  $([math]::Round((((750 * 8) / $avgUp)/60),2)) minutes  =  $([math]::Round(((((750 * 8) / $avgUp)/60)/60),2)) hours"
Write-Host "  1 GB:  $([math]::Round(((1000 * 8) / $avgUp),2)) seconds  =  $([math]::Round((((1000 * 8) / $avgUp)/60),2)) minutes  =  $([math]::Round(((((1000 * 8) / $avgUp)/60)/60),2)) hours"
Write-Host ""
"Test Date: $date" | Out-File -FilePath "C:\temp\$($env:COMPUTERNAME)_CadLinkSpeedTests.txt" -Append
"Test Duration: $timer seconds" | Out-File -FilePath "C:\temp\$($env:COMPUTERNAME)_CadLinkSpeedTests.txt" -Append
"Client sync batch size: $(($batchSize / 1024 / 1024)) MB" | Out-File -FilePath "C:\temp\$($env:COMPUTERNAME)_CadLinkSpeedTests.txt" -Append
"Average Latency: $averageLatency ms" | Out-File -FilePath "C:\temp\$($env:COMPUTERNAME)_CadLinkSpeedTests.txt" -Append
"Average Upload: $avgUp mpbs" | Out-File -FilePath "C:\temp\$($env:COMPUTERNAME)_CadLinkSpeedTests.txt" -Append
"Average Download $avgDown mpbs" | Out-File -FilePath "C:\temp\$($env:COMPUTERNAME)_CadLinkSpeedTests.txt" -Append
"" | Out-File -FilePath "C:\temp\$($env:COMPUTERNAME)_CadLinkSpeedTests.txt" -Append

pause
