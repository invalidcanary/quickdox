### Author: @InvalidCanary
### Purpose: grabs a handful of raw data, drops into an html file for ease of transport
### Parameters: no mandatory params
### Variables:  set $outputLocation to an appropriate spot
### Example: .\quickdox.ps1 
### Note: nothing fancy here, just raw dumps.  no error handling, no version checking.
### Note: Will throw warning on 2016 servers due to get-clientaccessServer deprecation, but is best for compat with 2013 right now because I'm a lazy scripter

### Future Feature Adds
### Add: export/auto import of cert
### Add: .NET version installed
### Add: mbx sizing
### Add: resource mailboxes
### Add: O365/hybrid user creation script/process/executable
### Add: Exchange supportability matrix (.net, CU, OS version, AD, etc), auth mechanisms
### Add: AADC reporting script (pull settings periodically for reference/change management)
### Add: RAM, cpu count
### Add: Intune policy enumeration
### Add: Conditional Access policy enumeration/healthcheck
### Add: which cert is active for IIS per server?

### Assign output location and create path if necessary

Param(
[Parameter(Mandatory=$false)]
[switch]$Remote,
[Parameter(Mandatory=$false)]
[string]$connectServer
)

$outputDir = "C:\scripts\quickdox\"
$outputFileName = "quickdox.htm"

$outputLocation = $outputDir + $outputFileName

If(!(test-path $outputDir))
{
      New-Item -ItemType Directory -Force -Path $outputDir
}

$connectionURL="http://" + $connectServer + "/PowerShell/"

### Connect to remote server if Remote parameter equals $true
If($Remote) {
### Ask for credential, build server connection, import session
    $cred = Get-Credential
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $connectionURL -Authentication Kerberos -Credential $cred
    Import-PSSession $session -DisableNameChecking
}


$Header = @"

<style>
BODY{font-family: Verdana; font-size: 10pt;}
H1{font-size: 22px;}
H2{font-size: 20px; padding-top: 10px;}
H3{font-size: 12px; padding-top: 8px;}
H4{font-size: 10px; padding-top: 8px;}
TABLE {font-family: verdana; font-size: 10px; border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH {border-width: 1px;padding: 3px;border-style: solid;border-color: gray;background-color: #e89d3a;}
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: gray;}
</style>
<title>
Exchange QuickDox
</title>

"@

$PowerHeader = @"
 <H3>Power Plan Settings</H3>
"@

$IPHeader = @"
 <H3>IP Address Settings</H3>
"@

$Pre = "<H2>Exchange Environment Data</H2>"
$Post = '<h4>Quick and Dirty report.  Yell at <a href="mailto:rvogsland@presidio.com?Subject=Exchange%20QuickDox">rvogsland@presidio.com</a> if you need something else routinely.</h4>'

$orgConfig = Get-organizationConfig | Select-Object -Property name, *mapi* | ConvertTo-Html -Fragment -precontent "$Pre <H3>Org Config</H3>"
$adfsAuth = Get-organizationConfig | Select-Object -Property name,  @{Name='AdfsAudienceUris';Expression={$_.servers -join ", "}}

If($adfsAuth.AdfsAudienceUris -ne "") {$adfsAuth.AdfsAudienceUris = "Enabled"} else {$adfsAuth.AdfsAudienceUris = "Disabled"}
$adfsAuth =$adfsAuth | ConvertTo-Html -Fragment -precontent "<H3>ADFS Auth</H3>"
$serverInfo = Get-ExchangeServer | Select-Object -Property name, serverrole, admindisplayversion, edition | ConvertTo-Html -Fragment -precontent "<H3>Server Information</H3>"


### get server IP information and Power Information
$exServers = get-exchangeServer | ForEach-Object { $_.Identity}

foreach ($Computer in $exServers) { 
$planSetting = Get-WmiObject -Class win32_powerplan -Namespace root\cimv2\power -Filter "isActive='true'" -ComputerName $Computer -ErrorAction SilentlyContinue
$planSetting | Select-Object @{Name="ComputerName"; Expression = {$_.PSComputerName}}, ElementName

$OutputObjPower  = New-Object -Type PSObject            
$OutputObjPower | Add-Member -MemberType NoteProperty -Name ComputerName -Value $planSetting.PSComputerName 
$OutputObjPower | Add-Member -MemberType NoteProperty -Name PowerPlan -Value $planSetting.ElementName

$PowerPlaninfo += $OutputObjPower | ConvertTo-Html -Fragment
}

     
 foreach ($Computer in $exServers) {            
  if(Test-Connection -ComputerName $Computer -Count 1 -ea 0) {            
   try {            
    $Networks = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $Computer -EA Stop | ? {$_.IPEnabled}            
   } catch {            
        Write-Warning "Error occurred while querying $computer."            
        Continue            
   }            
   foreach ($Network in $Networks) {            
    $IPAddress  = $Network.IpAddress[0]            
    $SubnetMask  = $Network.IPSubnet[0]            
    $DefaultGateway = "$($Network.DefaultIPGateway)"          
    $DNSServers  = "$($Network.DNSServerSearchOrder)"         
    $IsDHCPEnabled = $false            
    If($network.DHCPEnabled) {            
     $IsDHCPEnabled = $true            
    }            
    $MACAddress  = $Network.MACAddress            
    $OutputObj  = New-Object -Type PSObject            
    $OutputObj | Add-Member -MemberType NoteProperty -Name ComputerName -Value $Computer            
    $OutputObj | Add-Member -MemberType NoteProperty -Name IPAddress -Value $IPAddress            
    $OutputObj | Add-Member -MemberType NoteProperty -Name SubnetMask -Value $SubnetMask            
    $OutputObj | Add-Member -MemberType NoteProperty -Name Gateway -Value $DefaultGateway            
    $OutputObj | Add-Member -MemberType NoteProperty -Name IsDHCPEnabled -Value $IsDHCPEnabled            
    $OutputObj | Add-Member -MemberType NoteProperty -Name DNSServers -Value $DNSServers 
    # $OutputObj | Add-Member -MemberType NoteProperty -Name MACAddress -Value $MACAddress 

    # @{Name='DNSServerSearchOrder';Expression={$_.DNSServerSearchOrder -join '; '}}          
   # $OutputObj | Add-Member -MemberType NoteProperty -Name MACAddress -Value $MACAddress            
    
    If($DNSServers -ne ""){ ###after setting to goofy variable format above, $null no longer works as a filter and I'm lazy
    #$OutputObj
    $IPinfo += $OutputObj | ConvertTo-Html -Fragment 
    #-precontent "<H3>IP Information</H3>"
    }          
   }            
  }            
 }   
### end IP get



$Databases = Get-mailboxdatabase | Select-Object -Property AdminDisplayName, MasterType, ServerName | ConvertTo-Html -Fragment -precontent "<H3>Database Information</H3>"
$replStatus = Get-mailboxdatabasecopystatus * | Select-Object -Property Name, Status, CopyQueueLength, ReplayQueueLength, LastInspectedLogTime, ContentIndexState | ConvertTo-Html -Fragment -precontent "<H3>Database Replication Health</H3>"
$mbxCount = ((Get-Mailbox -ResultSize Unlimited).count).toString()
$mbxTitle = "Mailbox Count"
$mbxHTML = ConvertTo-Html -Fragment -precontent "<H3>User Mailbox Count</H3><table><th>$mbxTitle</th><tr><td>$mbxCount</td></tr></table>"
$owaVDirs = Get-OwaVirtualDirectory -ADPropertiesOnly | Select-Object -Property identity, *nalurl, *authentication, @{Name='internalauthenticationmethods';Expression={$_.internalauthenticationmethods -join ", "}}, @{Name='externalauthenticationmethods';Expression={$_.externalauthenticationmethods -join ", "}}| ConvertTo-Html -Fragment -precontent "<H3>OWA URLs</H3>"
$ecpVDirs = Get-ecpVirtualDirectory -ADPropertiesOnly | Select-Object -Property identity, *nalurl, *authentication, @{Name='internalauthenticationmethods';Expression={$_.internalauthenticationmethods -join ", "}}, @{Name='externalauthenticationmethods';Expression={$_.externalauthenticationmethods -join ", "}} | ConvertTo-Html -Fragment -precontent "<H3>ECP URLs</H3>"
$ewsVDirs = Get-webservicesVirtualDirectory -ADPropertiesOnly | Select-Object -Property identity, *nalurl, @{Name='internalauthenticationmethods';Expression={$_.internalauthenticationmethods -join ", "}}, @{Name='externalauthenticationmethods';Expression={$_.externalauthenticationmethods -join ", "}}, *authentication | ConvertTo-Html -Fragment -precontent "<H3>EWS URLs</H3>"
$easVDirs = Get-activesyncVirtualDirectory -ADPropertiesOnly | Select-Object -Property identity, *nalurl, @{Name='internalauthenticationmethods';Expression={$_.internalauthenticationmethods -join ", "}}, @{Name='externalauthenticationmethods';Expression={$_.externalauthenticationmethods -join ", "}} | ConvertTo-Html -Fragment -precontent "<H3>ActiveSync URLs</H3>"
$oabVDirs = Get-oabVirtualDirectory -ADPropertiesOnly | Select-Object -Property identity, *nalurl, *authentication, @{Name='internalauthenticationmethods';Expression={$_.internalauthenticationmethods -join ", "}}, @{Name='externalauthenticationmethods';Expression={$_.externalauthenticationmethods -join ", "}} | ConvertTo-Html -Fragment -precontent "<H3>OAB URLs</H3>"
$mapiVDirs = Get-mapiVirtualDirectory -ADPropertiesOnly | Select-Object -Property identity, *nalurl, @{Name='iisauthenticationmethods';Expression={$_.iisauthenticationmethods -join ", "}}, @{Name='internalauthenticationmethods';Expression={$_.internalauthenticationmethods -join ", "}}, @{Name='externalauthenticationmethods';Expression={$_.externalauthenticationmethods -join ", "}} | ConvertTo-Html -Fragment -precontent "<H3>MAPI URLs</H3>"
$autodiscoURIs = get-clientaccessService | Select-Object -Property identity, AutoDiscoverServiceInternalUri | ConvertTo-Html -Fragment -precontent "<H3>AutoDiscover Configuration URLs</H3>"
$OAconfig = Get-OutlookAnywhere -ADPropertiesOnly | Select-Object -Property server, *nalhostname, *clientauth*, *ssl*, @{Name='iisauthenticationmethods';Expression={$_.iisauthenticationmethods -join ", "}} | ConvertTo-Html -Fragment -precontent "<H3>Outlook Anywhere Hosts</H3><p>"
$DAGconfig = Get-DatabaseAvailabilityGroup | Select-Object -Property name, @{Name='servers';Expression={$_.servers -join ", "}} , *centeract*, exchangeversion, @{Name='DatabaseAvailabilityGroupIpv4Addresses';Expression={$_.DatabaseAvailabilityGroupIpv4Addresses -join ", "}}, witness* | ConvertTo-Html -Fragment -precontent "<H3>DAG Information</H3>"

ConvertTo-Html -Head $Header -postcontent $Post -Body "$orgConfig $adfsAuth $DAGconfig $serverInfo $Databases $replStatus $mbxHTML $owaVDirs $ecpVDirs $ewsVDirs $easVDirs $oabVDirs $mapiVDirs $autodiscoURIs $OAconfig $IPHeader $IPinfo $PowerHeader $PowerPlaninfo" -title "Exchange QuickDox" | out-file $outputLocation
#if send=yes then blah blah
#Send-MailMessage -Attachments $outputLocation -From Exchange@invalidcanary.com -To randallv@dawho.com -Subject "QuickDox" -Body "output of quickdox script." -SmtpServer phobos.dawho.com -credential $cred


# @{Name='iisauthenticationmethods';Expression={($_ | Select -ExpandProperty iisauthenticationmethods | Select -ExpandProperty Name) -join ","}}
# @{Name='DatabaseAvailabilityGroupIpv4Addresses';Expression={$_.DatabaseAvailabilityGroupIpv4Addresses -join ", "}} 

