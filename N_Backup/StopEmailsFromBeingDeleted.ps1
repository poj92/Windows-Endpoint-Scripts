Set-ExecutionPolicy RemoteSigned
Install-Module -Name ExchangeOnlineManagement
Connect-ExchangeOnline


LIFT-Events@lift-financial.com

Set-CalendarProcessing -Identity "LIFT-Events@lift-financial.com" -DeleteNonCalendarItems $false