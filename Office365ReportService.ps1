#$ErrorActionPreference = "SilentlyContinue"
if ($MyInvocation.MyCommand.CommandType -eq "ExternalScript") {
    $presentpath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
}
else {
    $presentpath = Split-Path -Parent -Path ([Environment]::GetCommandLineArgs()[0])
}
$signature = @"
[DllImport(@"$presentpath\Support\EtsIns.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
extern public static int CollectTrapsForSending(
    StringBuilder srcSystem,
    StringBuilder pSystems,
    StringBuilder pPorts,
    int pLogType,
    int pEvtType,
    StringBuilder pEvtSource,
    int pEvtCategory,
    int pEvtId,
	int pticks,
    StringBuilder pEvtDesc,
	StringBuilder UserName,
	StringBuilder Domain);
"@
Add-Type -MemberDefinition $signature -Name SendTrap -Namespace CollectTrapsForSending -Using System.Text -PassThru | Out-Null
$signature = @"
[DllImport(@"$presentpath\Support\EtsIns.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
extern public static int StartSendTCPTrapsEx(
    StringBuilder pSystems,
    StringBuilder pPorts,
    StringBuilder srcSystem,
    StringBuilder DestIPAddress,	
	int pDelay);
"@
Add-Type -MemberDefinition $signature -Name SendTrap -Namespace StartSendTCPTrapsEx -Using System.Text -PassThru | Out-Null
$signature = @"
[DllImport(@"$presentpath\Support\EtsIns.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
extern public static int StartSendUDPTrapsEx(
    StringBuilder pSystems,
    StringBuilder pPorts,
    StringBuilder srcSystem,
    StringBuilder DestIPAddress,	
	int pDelay);
"@
Add-Type -MemberDefinition $signature -Name SendTrap -Namespace StartSendUDPTrapsEx -Using System.Text -PassThru | Out-Null
Add-Type -AssemblyName SYSTEM.WEB
$Details = Import-Clixml $presentpath\Support\Details.xml
$credential = new-object -typename System.Management.Automation.PSCredential -argumentlist ($Details.OfficeUser, ($Details.OfficePass | ConvertTo-SecureString))
function Fetch-OfficeLog ($LogType, $Delay) {
    try {
        $lastpick = Import-Clixml $presentpath\Support\LastPick-$LogType.xml
        $starttime = Get-Date $lastpick.startDate -Format "yyyy-MM-ddTHH:mm:ss"
        if (!$starttime) {
            $starttime = Get-Date (Get-Date).AddMinutes(-60 - $Delay).ToUniversalTime() -Format "yyyy-MM-ddTHH:mm:ss"
        }
    }
    catch {
        if ($_.CategoryInfo.Reason -eq "FileNotFoundException") {
            $starttime = Get-Date (Get-Date).AddMinutes(-60 - $Delay).ToUniversalTime() -Format "yyyy-MM-ddTHH:mm:ss"
        }
    }
    $endtime = Get-Date (Get-Date).AddMinutes(-$Delay).ToUniversalTime() -Format "yyyy-MM-ddTHH:mm:ss"
    $uri = "https://reports.office365.com/ecp/reportingwebservice/reporting.svc/$LogType`?`$filter=StartDate%20eq%20datetime'$starttime'%20and%20EndDate%20eq%20datetime'$endtime'"
    do {
        (Invoke-RestMethod -uri ($uri + "&`$format=Json") -Credential $credential) | % {
            $uri = $_.d.__next
            $field = (($_.d.results | select -Last 1).psobject.properties | Where-Object {($_.name -eq "Received") -or (($_.name -eq "Date"))}).name
            [pscustomobject]@{
                StartDate = $endtime
            } | Export-Clixml $presentPath\Support\LastPick-$LogType.xml
            $_.d.results |Select-Object -Property *, @{N = 'Type'; E = {$_.__metadata.type}} -ExcludeProperty startdate, enddate, index, __metadata | % {
                [CollectTrapsForSending.SendTrap]::CollectTrapsForSending($Details.Organisation, $Details.'ET Manager',$Details.'ET Manager Port', 3, 3, $LogType, 0, 3230, (get-date $_.$field -UFormat "%s"), ($_ | Out-String).Trim(), "NA","NA") | Out-Null
            }
            if($details.protocol -eq 'TCP'){
            [StartSendTCPTrapsEx.SendTrap]::StartSendTCPTrapsEx($Details.'ET Manager',$Details.'ET Manager Port',$Details.Organisation,(Resolve-DnsName -Name $Details.'ET Manager' -Type A).IPAddress,0) | out-null
            } else{
            [StartSendUDPTrapsEx.SendTrap]::StartSendUDPTrapsEx($Details.'ET Manager',$Details.'ET Manager Port',$Details.Organisation,(Resolve-DnsName -Name $Details.'ET Manager' -Type A).IPAddress,0) | out-null
            }
        }
    } until ($uri -eq $null)
  }
while ($true){
    Fetch-OfficeLog -LogType "MessageTrace" -Delay 10
    Start-Sleep -Seconds 5
}
$queryresult = Invoke-RestMethod -Uri "https://reports.office365.com/ecp/reportingwebservice/reporting.svc/MessageTraceDetail?`$filter=MessageTraceId%20eq%20guid'170c9dd0-75aa-4431-16a5-08d652d53e23'%20and%20RecipientAddress%20eq%20'essentials-support@eventtracker.com'%20and%20SenderAddress%20eq%20'MDR50@EVENTTRACKER.COM'&`$format=Json" -Credential $credential