#-----------[Script Commenting]----------
# ---------------------------------------------------------------------------
# FILE           	:	Identifying_unrestricted_RDP_access.ps1  
# DESCRIPTION       : 	Identifying Security groups whcih have unrestricted RDP access. It will replace the public route (0.0.0.0/0) with HARMAN IP ranges(10.0.0.0/8 & 172.16.0.0/12)
# AUTHOR         	:	everestdx@everestdx.com 
# COMPANY           : 	EverestDXi 
# VERSION        	:	1.0
# CREATED        	: 	03.26.2019
# REVISION       	: 	
# NOTES          	:     
# Contact Person    : 	Mohit Agrawal
#---------------------------------------------------------------------------
#---------Declarations----------
$BaseName=$null
$Timer = [Diagnostics.Stopwatch]::StartNew()
$location = (get-location).path
$predirectory = (Resolve-Path (Join-Path $location '..')).path
$emailTemplate = "$location\Config\AWS_Email.html"
$local  = Get-Content -Path $location\Config\localconfig.json | ConvertFrom-Json
$global = Get-Content -Path $predirectory\globalconfig.json | Convertfrom-Json
$fdate = Get-Date -Format "yyyyMMdd"
$BaseName = (Get-Item $MyInvocation.MyCommand.Name| Select-Object -ExpandProperty BaseName)
$logfile  = $location+"\Logs\"+$BaseName+"_"+$fdate+".log"
$Summary = $null
Get-ChildItem "$location\Logs\" -Recurse -File | Where CreationTime -lt  (Get-Date).AddDays(-30)  | Remove-Item -Force
if (!(Test-Path $logfile))
{
    New-Item -ItemType "file" -Path $logfile
}
$MailMessage = $null
$mailobj = $null
$mailbody = $null
$MailMessage =@()
$mailbody = @()
$Summary = @()
if ($local.region -ne $null) {
    $allregions = $local.region
}
else {
    $allregions = $global.allregions
}
if ($local.profilenames -ne $null) {
    $allprofiles = $local.profilenames
}
else {
    $allprofiles = $global.profilenames
}
$bodyheader = $local.report
$MailMessage = $null

#---------[Functions]-----------

Function sendEmail($from, $to, $subject, $body,$att) {
    $smtpprops=$global
    $useSES=$false
    if($smtpprops.smtp.server -ne '') {
        if($from -eq ''){
            $from=$smtpprops.defaultFrom
        }
    $mail = New-Object System.Net.Mail.Mailmessage
    $mail.IsBodyHTML=$true
    $mail.from = $from
    if($to.GetType().BaseType.Name -eq 'Array') {
        foreach($email in $to){
            $mail.to.add($email)
        }
    }
    else{ 
        $mail.to.add($to)
    }
    if($att -ne ''){
        $mail.Attachments.Add($att)
    }
    $mail.subject = $subject
    $mail.body = $body
    $server = $smtpprops.smtp.Server
    $port = $smtpprops.smtp.port
    $Smtp = New-Object Net.Mail.SmtpClient($server, $port)
    $Smtp.EnableSsl = $false
    if($pswd -ne '' -and $userid -ne ''){
        $Smtp.Credentials = New-Object System.Net.NetworkCredential($userid, $pswd)
        $userid = $smtpprops.smtp.userid
        $pswd = $smtpprops.smtp.secret
        try{
            $Smtp.send($mail)
        }
        catch{
            write-host "Unable to send email: " $_
            $useSES=$true
        }
    }
    else{
        $useSES=$true
    }
    if($useSES){
        Send-SESEmail -Source $from -Destination_ToAddress $to -Html_Data $body -Subject_Data $subject}
    }
}

#---------[Log Function]-----------

Function Write-Log($Source,$Message){
    $thedate = Get-Date -Format "MM-dd-yyyy"
    $thedate = $thedate.ToString()
    (get-date -format "MM-dd-yyyy HH:mm:ss") + " [" + $source + "] " + $message >> "$location\Logs\$thedate.txt"
}

#---------[Main Script]--------
try{
    Write-Log " INFO ","Script Started ...."       

foreach ($prof in $allprofiles ) {
        $profile = $prof.profilename
        $AccountName = $prof.account

        if( $local.debug -eq 0 ) {Write-Log " DEBUG ", "Using Profile - $profilename"}
        Write-Log " INFO ", "Using Profile - $profile"

        foreach($region in $allregions) {

            Write-Log " INFO  - Region... ", $region
            $MailMessage = '<table><tr><th>Account</th><th>Region</th><th>Volume ID</th><th>Volume Type</th><th>Volume IOPS</th><th>Volume Size</th><th>Volume Status</th><th>Volume SnapshotID</th><th>Attached EC2 with Volume</th>'
            $Volumes =  Get-EC2Volume -Region $Region
          
            foreach($VolumeId in ($Volumes)) {

           
                $mailobj = new-object PSObject -Property @{
                                            AccountName = $AccountName
                                            Region = $region                                            
                                             VolumeId = ($VolumeId).VolumeId
                    Volume_Type = (($VolumeId).VolumeType).Value
                    IOPS = ($VolumeId).Iops
                    Volume_Size = ($VolumeId).Size
                    Attached_EC2Instance = (($VolumeId).Attachments).InstanceId
                    Volume_Status = (($VolumeId).State).Value
                    SnapshotId = ($VolumeId).SnapshotId
                    
                    
                                       
                } 
                Write-Log $Source ( "Account: "+$AccountName+ "Region: "+$region+"VolumeId: "+($Volumes).VolumeId+"Volume_Type: "+($VolumeId).VolumeType+"IOPS: "+($Volumes).Iops+"Volume_Size: "+($Volumes).Size+"Volume_Status: "+($Volumes).State+"SnapshotId: "+($Volumes).SnapshotId+"Attached_EC2Instance: "+(($VolumeId).Attachments).InstanceId)
                $mailbody += $mailobj 
            }
        }
    } 

   
   if($mailbody -ne ""){
        foreach($mail in ($mailbody | Group-Object -Property "AccountName")){ 
            $AccountCount = $mail.Count
            $MailMessage += '<tr><td style="vertical-align : top;" rowspan='+$AccountCount+'>'+$mail.Name+'</td>' 
            foreach($reg in ($mailbody | Where-Object {$_.AccountName -eq $mail.Name} | Group-Object -Property 'Region')){
                $regcount = $reg.count
                $MailMessage += '<td style="vertical-align : top;" rowspan='+$regCount+'>'+$reg.Name+'</td>'
                foreach($inst in $reg.Group){
                    $MailMessage += '<td>'+ ($inst).VolumeId +'</td><td>' + ($inst).Volume_Type + '</td><td>' + ($inst).IOPS + '</td><td>' + ($inst).Volume_Size + '</td><td>' + ($inst).Volume_Status + '</td><td>' + ($inst).SnapshotId + '</td><td>' + ($inst).Attached_EC2Instance + '</td></tr>'
                }  
            }
        }
    }

$colspa = ($mailbody | Group-Object -Property "AccountName").length
if ($colspa -le 3){
$colspan = $colspa
}
else{
$colspan = 3
}

$Summary = '<table class="table1"><tr><td colspan='+$colspan+'><font size="5" color="#000" weight="bold"><b>'+$local.summarytitle+' </b></font><font size="5" color="#FF4500"><b>'+$mailbody.Count+'</b></font></td></tr><tr>'

foreach($acco in ($mailbody | Group-Object -Property "AccountName")){
$Summary += '<td><font size="3">'+$acco.Name+'</font>: <font size="3" color="#FF7900">'+$acco.Count+'</font><br/>'
foreach($regio in ($mailbody | Where-Object {$_.AccountName -eq $acco.Name} | Group-Object -Property 'Region')){
$Summary += $regio.Name+': <font size="3" color="#FF7900">'+$regio.Count+'</font><br/>'
}
$Summary += '</td>'
if($i -eq 2){
$Summary += '</tr><tr>'
$i=0
}
}
$Summary += '</tr></table><br/>'
$MailMessage += '</table>'

if($mailbody -ne ""){
$htmltemplate = [System.IO.File]::ReadAllText($emailTemplate).Replace("#Content",$MailMessage).Replace("#reportheader",$local.reportheader).Replace("#Summary",$Summary).Replace("#pagetitle", $local.pagetitle).Replace("#maintitle", $global.maintitle)
$htmltemplate > $location"/Reports/"$BaseName"_Report.html"
$from = ''
$to = $local.toaddress
$subject = $local.subject
$body = $htmltemplate
#sendEmail $from $to $subject $body '' 
Write-Log " INFO ","Script End ...."
        if($local.debug -eq 0 ) {Write-Log " DEBUG ","End of the script"}
}
}
catch
{
$ErrorMessage = $_
write-log " ERROR ",$ErrorMessage
if($local.debug -eq '0' ) {Write-Log " DEBUG ", $ErrorMessage}
}
$Timer.Stop()

write-log " INFO ",("Elapsed time to run Script : "+ [string]$Timer.elapsed)

 
