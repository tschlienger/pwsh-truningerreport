Start-Transcript "lastexec.log"

$DeviceIp = $args[0]
$DeviceIpFormatted = $DeviceIp -replace "\.", "-"
$DateFormat = "dd\/MM\/yyyy"
$ExportFile = "Daten_$DeviceIpFormatted.xml"
$ApiUrlTemplate = "http://{0}/hp/device/webAccess/open_{1}.xml?startDate={2}&endDate={3}"
$MonthNames = @("Januar","Februar","März","April","Mai","Juni","Juli","August","September","Oktober","November","Dezember")
$MailRecipient = @("itsource@truninger-plot24.ch")
$MailSender = ""
$MailServer = ""
$MailPort = "25"
$MailBody = "Guten Tag, im Anhang finden Sie den Zählerstandsbericht des vergangenen Monats.`r`n`r`nDies ist ein automatisch generierter Bericht."

$Date = Get-Date
$DaysOfLastMonth = $Date.AddDays(-$Date.Day).Day
$StartDate = Get-Date $Date.AddDays(-($DaysOfLastMonth+$Date.Day-1))
$StartDateString = Get-Date $StartDate -Format $DateFormat
$EndDate = Get-Date $Date.AddDays(-$Date.Day)
$EndDateString = Get-Date $EndDate -Format $DateFormat

$AccountingUrl = $ApiUrlTemplate -f $DeviceIp, "accounting", $StartDateString, $EndDateString
$UsageUrl = $ApiUrlTemplate -f $DeviceIp, "usage", $StartDateString, $EndDateString

$Accounting = Invoke-RestMethod -Method Get -Uri $AccountingUrl
$Usage = Invoke-RestMethod -Method Get -Uri $UsageUrl

$ExportXml = "<?xml version=`"1.0`" encoding=`"UTF-8`"?>`r`n<AccountInfo version=`"1.2`">`r`n<Current_Printer_Configuration>`r`n$($Accounting.Serviceability.Current_Printer_Configuration.InnerXml)`r`n</Current_Printer_Configuration>`r`n$($Accounting.Serviceability.JOBS_ACCOUNTING_INFO.InnerXml -replace " />", "/>")`r`n<Printer_Usage>`r`n$($Usage.Serviceability.Printer_Usage.InnerXml -replace " />", "/>")`r`n</Printer_Usage>`r`n</AccountInfo>"

Set-Content $ExportFile -Encoding UTF8 -Value $ExportXml
Send-MailMessage -Subject "Zählerstandsmeldung $($MonthNames[$StartDate.Month-1]) $($StartDate.Year)" -From $MailSender -To $MailRecipient -SmtpServer $MailServer -Port $MailPort -Attachments $ExportFile -Encoding utf8 -Body $MailBody

Stop-Transcript