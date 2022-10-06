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

$date = Get-Date -Format "yyyy/MM/dd HH:mm:ss"
$logresults = 300
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

Write-Host "Initiated at $date"
Write-Host "Running speed tests for 5 minutes"
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

Write-Host "Speed Tests complete. Test Duration: $timer seconds"

if (($averageLatency -gt 0) -and ($averageLatency -le 20)) {
    
    Write-Host "Low Average Latency: $averageLatency ms" #-ForegroundColor Green

} elseif (($averageLatency -gt 20) -and ($averageLatency -le 50)) {
    
    Write-Host "Medium Latency: $averageLatency ms" #-ForegroundColor Yellow

} elseif (($averageLatency -gt 50) -and ($averageLatency -le 75)) {
    
    Write-Host "High Average Latency: $averageLatency ms" #-ForegroundColor Red

} elseif ($averageLatency -gt 75) {

    Write-Host "Very High Average Latency: $averageLatency ms" #-ForegroundColor Red    

}

Write-Host "Average Upload Speed: $avgUp mbps"
Write-Host "Average Download Speed: $avgDown mbps"
Write-Host ""
"Test Date: $date" | Out-File -FilePath "C:\temp\$($env:COMPUTERNAME)_CadLinkSpeedTests.txt" -Append
"Test Duration: $timer seconds" | Out-File -FilePath "C:\temp\$($env:COMPUTERNAME)_CadLinkSpeedTests.txt" -Append
"Average Latency: $averageLatency ms" | Out-File -FilePath "C:\temp\$($env:COMPUTERNAME)_CadLinkSpeedTests.txt" -Append
"Average Upload: $avgUp mpbs" | Out-File -FilePath "C:\temp\$($env:COMPUTERNAME)_CadLinkSpeedTests.txt" -Append
"Average Download $avgDown mpbs" | Out-File -FilePath "C:\temp\$($env:COMPUTERNAME)_CadLinkSpeedTests.txt" -Append
"" | Out-File -FilePath "C:\temp\$($env:COMPUTERNAME)_CadLinkSpeedTests.txt" -Append

pause
