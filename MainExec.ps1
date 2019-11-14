#Use Static path for Task Scheduler
#$fpath = "C:\inetpub\wwwroot\healthreports"

#Use Local path for testing scripts
$fpath = "."

Write-Host "Initizalizing System Report."
Write-Host "-------------------------------"
Write-Host "This may take a while."
#Run system report script.
.("$fpath\GetSystemReport.ps1")

Write-Host "Done."
Write-Host "-------------------------------"

Write-Host "Writing Reports to Collection."
Write-Host "-------------------------------"
#Add Files to Text Document.

Get-childitem -path "$fpath\Reports\" -rec -file | Sort-Object LastWriteTime -Descending | select-object -expandproperty name | out-file "$fpath\FileList.txt"

Write-Host "Done."
Write-Host "-------------------------------"

Write-Host "Updating Reports Page."
Write-Host "-------------------------------"
#Run HTML script.
.("$fpath\CreateHTML.ps1")

Write-Host "-------------------------------"
Write-Host "Done."
Write-Host "-------------------------------"
Write-Host "Sending Email to IT Group."
Write-Host "-------------------------------"

$mailParams = @{
    to         = 'IT <it@YourDomain.com>'
    from       = 'IT <it@YourDomain.com>'
    subject    = 'System Health Report for ' + ($(Get-Date -Format "MM/dd/yyyy"))
    smtpserver = 'ExchangeServer.domain.local'
    body       = 'System Health Reports have completed: <a href="http://IIS-URL-or-IP/healthreports/">View Reports</a>'
}
Send-MailMessage @mailParams -BodyAsHtml

Write-Host "Done."
Write-Host "-------------------------------"
start-sleep -s 60