<#
  DESCRIPTION 
   Script to Automated Email Reminders when Users Passwords due to Expire. 
   Original by Robert Pearman can be found here: https://web.archive.org/web/20161116225450/https://gallery.technet.microsoft.com/Password-Expiry-Email-177c3e27
   This version is made to read the configs or flags for this script from a JSON file that is in the same folder. This was created as an experiment and it function.
   Alexis Daigle.
 #>
 
# Import Configurations from the JSON Config File
$configPath = "$PSScriptRoot\PasswordChangeConfig.json"
try{ 
$Conf = Get-Content -Path $configPath | ConvertFrom-Json
}
catch{ 
    Write-Warning "Unable to load Config file @ $configPath" 
} 
[string]$smtpServer = $conf.smtpServer
[int]$expireInDays = $conf.expireInDays
[string]$from = $conf.from
[bool]$logging = if ($conf.Logging -like "true") { $true } else { $false }
[string]$logPath = $conf.logpath
[bool]$testing = if ($conf.testing -like "true") { $true } else { $false }
[string]$testRecipient = $conf.testRecipient
[bool]$status = if ($conf.status -like "true") { $true } else { $false }
[string]$reportto = $conf.reportto
[array]$interval = $conf.interval

################################################################################################################### 
# Time / Date Info 
$start = [datetime]::Now 
$midnight = $start.Date.AddDays(1) 
$timeToMidnight = New-TimeSpan -Start $start -end $midnight.Date 
$midnight2 = $start.Date.AddDays(2) 
$timeToMidnight2 = New-TimeSpan -Start $start -end $midnight2.Date 
# System Settings 
$textEncoding = [System.Text.Encoding]::utf8NoBOM 
$today = $start 
# End System Settings 
 
# Load AD Module 
try{ 
    Import-Module ActiveDirectory -ErrorAction Stop 
} 
catch{ 
    Write-Warning "Unable to load Active Directory PowerShell Module" 
} 
# Set Output Formatting - Padding characters 
$padVal = "20" 
Write-Output "Script Loaded" 
Write-Output "*** Settings Summary ***" 
$smtpServerLabel = "SMTP Server".PadRight($padVal," ") 
$expireInDaysLabel = "Expire in Days".PadRight($padVal," ") 
$fromLabel = "From".PadRight($padVal," ") 
$testLabel = "Testing".PadRight($padVal," ") 
$testRecipientLabel = "Test Recipient".PadRight($padVal," ") 
$logLabel = "Logging".PadRight($padVal," ") 
$logPathLabel = "Log Path".PadRight($padVal," ") 
$reportToLabel = "Report Recipient".PadRight($padVal," ") 
$interValLabel = "Intervals".PadRight($padval," ") 
$statusLabel = "status".PadRight($padval," ") 
# Testing Values 
if($testing) 
{ 
    if(($testRecipient) -eq $null) 
    { 
        Write-Output "No Test Recipient Specified" 
        Exit 
    } 
} 
# Logging Values 
if($logging) 
{ 
    if(($logPath) -eq $null) 
    { 
        $logPath = $PSScriptRoot 
    } 
} 
# Output Summary Information 
Write-Output "$smtpServerLabel : $smtpServer" 
Write-Output "$expireInDaysLabel : $expireInDays" 
Write-Output "$fromLabel : $from" 
Write-Output "$logLabel : $logging" 
Write-Output "$logPathLabel : $logPath" 
Write-Output "$testLabel : $testing" 
Write-Output "$testRecipientLabel : $testRecipient" 
Write-Output "$reportToLabel : $reportto" 
Write-Output "$interValLabel : $interval" 
Write-Output "$statusLabel : $status"  
Write-Output "*".PadRight(25,"*") 
# Get Users From AD who are Enabled, Passwords Expire and are Not Currently Expired 
# To target a specific OU - use the -searchBase Parameter -https://docs.microsoft.com/en-us/powershell/module/addsadministration/get-aduser 
# You can target specific group members using Get-AdGroupMember, explained here https://www.youtube.com/watch?v=4CX9qMcECVQ  
# based on earlier version but method still works here. 
$users = get-aduser -filter {(Enabled -eq $true) -and (PasswordNeverExpires -eq $false)} -properties Name, PasswordNeverExpires, PasswordExpired, PasswordLastSet, EmailAddress | where { $_.passwordexpired -eq $false } 
# Count Users 
$usersCount = ($users | Measure-Object).Count 
Write-Output "Found $usersCount User Objects" 
# Collect Domain Password Policy Information 
$defaultMaxPasswordAge = (Get-ADDefaultDomainPasswordPolicy -ErrorAction Stop).MaxPasswordAge.Days  
Write-Output "Domain Default Password Age: $defaultMaxPasswordAge" 
# Collect Users 
$colUsers = @() 
# Process Each User for Password Expiry 
Write-Output "Process User Objects" 
foreach ($user in $users) 
{ 
    # Store User information 
    $Name = $user.Name 
    $emailaddress = $user.emailaddress 
    $passwordSetDate = $user.PasswordLastSet 
    $samAccountName = $user.SamAccountName 
    $pwdLastSet = $user.PasswordLastSet 
    # Check for Fine Grained Password 
    $maxPasswordAge = $defaultMaxPasswordAge 
    $PasswordPol = (Get-AduserResultantPasswordPolicy $user)  
    if (($PasswordPol) -ne $null) 
    { 
        $maxPasswordAge = ($PasswordPol).MaxPasswordAge.Days 
    } 
    # Create User Object 
    $userObj = New-Object System.Object 
    $expireson = $pwdLastSet.AddDays($maxPasswordAge) 
    $daysToExpire = New-TimeSpan -Start $today -End $Expireson 
    # Round Expiry Date Up or Down 
    if(($daysToExpire.Days -eq "0") -and ($daysToExpire.TotalHours -le $timeToMidnight.TotalHours)) 
    { 
        $userObj | Add-Member -Type NoteProperty -Name UserMessage -Value "today." 
    } 
    if(($daysToExpire.Days -eq "0") -and ($daysToExpire.TotalHours -gt $timeToMidnight.TotalHours) -or ($daysToExpire.Days -eq "1") -and ($daysToExpire.TotalHours -le $timeToMidnight2.TotalHours)) 
    { 
        $userObj | Add-Member -Type NoteProperty -Name UserMessage -Value "tomorrow." 
    } 
    if(($daysToExpire.Days -ge "1") -and ($daysToExpire.TotalHours -gt $timeToMidnight2.TotalHours)) 
    { 
        $days = $daysToExpire.TotalDays 
        $days = [math]::Round($days) 
        $userObj | Add-Member -Type NoteProperty -Name UserMessage -Value "in $days days." 
    } 
    $daysToExpire = [math]::Round($daysToExpire.TotalDays) 
    $userObj | Add-Member -Type NoteProperty -Name UserName -Value $samAccountName 
    $userObj | Add-Member -Type NoteProperty -Name Name -Value $Name 
    $userObj | Add-Member -Type NoteProperty -Name EmailAddress -Value $emailAddress 
    $userObj | Add-Member -Type NoteProperty -Name PasswordSet -Value $pwdLastSet 
    $userObj | Add-Member -Type NoteProperty -Name DaysToExpire -Value $daysToExpire 
    $userObj | Add-Member -Type NoteProperty -Name ExpiresOn -Value $expiresOn 
    # Add userObj to colusers array 
    $colUsers += $userObj 
} 
# Count Users 
$colUsersCount = ($colUsers | Measure-Object).Count 
Write-Output "$colusersCount Users processed" 
# Select Users to Notify 
$notifyUsers = $colUsers | where { $_.DaysToExpire -le $expireInDays} 
$notifiedUsers = @() 
$notifyCount = ($notifyUsers | Measure-Object).Count 
Write-Output "$notifyCount Users with expiring passwords within $expireInDays Days" 
# Process notifyusers 
foreach ($user in $notifyUsers) 
{ 
    # Email Address 
    $samAccountName = $user.UserName 
    $emailAddress = $user.EmailAddress 
    # Set Greeting Message 
    $name = $user.Name 
    $messageDays = $user.UserMessage 
    # Subject Setting 
    $subject="Your password will expire $messageDays" 
    # Email Body Set Here, Note You can use HTML, including Images. 
    # examples here https://youtu.be/iwvQ5tPqgW0  
    $body =" 
    <font face=""verdana""> 
    Hi $name, 
    <p><b> Your Password will expire $messageDays</b><br>
    </P>  
    To change your password on a PC press CTRL+ALT+Delete and choose Change Password. <br> 
    <p> If you are working remotly, make sure you are connected to the VPN while changing your password.
    </P> 
    <p>Thanks, <br>  
    </P> 
    HealthConnect IT 
    <a href=""mailto:ITEmail@ITCompany.Exp""?Subject=Password Expiry Assistance"">ITEmail@ITCompany.Exp</a>  
    </font>" 
    # If Testing Is Enabled - Email Administrator 
    if($testing) 
    { 
        $emailaddress = $testRecipient 
    } # End Testing 
    # If a user has no email address listed 
    if(($emailaddress) -eq $null) 
    { 
        $emailaddress = $testRecipient     
    }# End No Valid Email 
    $samLabel = $samAccountName.PadRight($padVal," ") 
    try{ 
        # If using interval paramter - follow this section 
        if($interval) 
        { 
            $daysToExpire = [int]$user.DaysToExpire 
            # check interval array for expiry days 
            if(($interval) -Contains($daysToExpire)) 
            { 
                # if using status - output information to console 
                if($status) 
                { 
                    Write-Output "Sending Email : $samLabel : $emailAddress" 
                } 
                # Send message
                Send-Mailmessage -smtpServer $smtpServer -from $from -to $emailaddress -subject $subject -body $body -bodyasHTML -priority High -ErrorAction Stop 
                $user | Add-Member -MemberType NoteProperty -Name SendMail -Value "OK" 
            } 
            else 
            { 
                # if using status - output information to console 
                # No Message sent 
                if($status) 
                { 
                    Write-Output "Sending Email : $samLabel : $emailAddress : Skipped - Interval" 
                } 
                $user | Add-Member -MemberType NoteProperty -Name SendMail -Value "Skipped - Interval" 
            } 
        } 
        else 
        { 
            # if not using interval paramter - follow this section 
            # if using status - output information to console 
            if($status) 
            { 
                Write-Output "Sending Email : $samLabel : $emailAddress" 
            } 
            Send-Mailmessage -smtpServer $smtpServer -from $from -to $emailaddress -subject $subject -body $body -bodyasHTML -priority High -ErrorAction Stop 
            $user | Add-Member -MemberType NoteProperty -Name SendMail -Value "OK" 
        } 
    } 
    catch{ 
        # error section 
        $errorMessage = $_.exception.Message 
        # if using status - output information to console 
        if($status) 
        { 
           $errorMessage 
        } 
        $user | Add-Member -MemberType NoteProperty -Name SendMail -Value $errorMessage     
    } 
    $notifiedUsers += $user 
} 
if($logging) 
{ 
    # Create Log File 
    Write-Output "Creating Log File" 
    $day = $today.Day 
    $month = $today.Month 
    $year = $today.Year 
    $date = "$day-$month-$year" 
    $logFileName = "$date-PasswordLog.csv" 
    if(($logPath.EndsWith("\"))) 
    { 
       $logPath = $logPath -Replace ".$" 
    } 
    $logFile = $logPath, $logFileName -join "\" 
    Write-Output "Log Output: $logfile" 
    $notifiedUsers | Export-CSV $logFile 
    if($reportTo) 
    { 
        $reportSubject = "Password Expiry Report" 
        $htmlhead = "<html>
				<style>
				BODY{font-family: Arial; font-size: 12pt;}
				H1{font-size: 22px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
				H2{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
				H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
				TABLE{border: 1px solid black; border-collapse: collapse; font-size: 11pt;}
				TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
				TD{border: 1px solid #969595; padding: 5px; }
				td.pass{background: #B7EB83;}
				td.warn{background: #FFF275;}
				td.fail{background: #FF2626; color: #ffffff;}
				td.info{background: #85D4FF;}
				</style>
				<body>
                <br>
                <b>Password Expiry For the next 21 Days</b><br>
                <br>"
        $htmltail = "<br></body></html>"
        $html = $notifiedUsers | select Name, EmailAddress, DaysToExpire, SendMail | ConvertTo-Html -Fragment
        $reportBody = $htmlhead + $html + $htmltail
        try{ 
            Send-Mailmessage -smtpServer $smtpServer -from $from -to $reportTo -subject $reportSubject -body $reportbody -bodyasHTML -Attachments $logFile -ErrorAction Stop  
        } 
        catch{ 
            $errorMessage = $_.Exception.Message 
            Write-Output $errorMessage 
        } 
    } 
} 
$notifiedUsers | select UserName,Name,EmailAddress,PasswordSet,DaysToExpire,ExpiresOn | sort DaystoExpire | FT -autoSize 
 
$stop = [datetime]::Now 
$runTime = New-TimeSpan $start $stop 
Write-Output "Script Runtime: $runtime" 
# End 
