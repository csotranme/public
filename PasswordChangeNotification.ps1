<#

.DESCRIPTION
   Script to Automated Email Reminders when Users Passwords due to Expire.

.EXAMPLE
  PasswordChangeNotification.ps1 -smtpServer mail.domain.com -expireInDays 21 -from "IT Support <support@domain.com>" -Logging -LogPath "c:\logFiles" -testing -testRecipient support@domain.com
  
  This example will use mail.domain.com as an smtp server, notify users whose password expires in less than 21 days, send mail from support@domain.com
  Logging is enabled, log path is c:\logfiles
  Testing is enabled, and test recipient is support@domain.com

.EXAMPLE
  PasswordChangeNotification.ps1 -smtpServer mail.domain.com -expireInDays 21 -from "IT Support <support@domain.com>" -reportTo myaddress@domain.com -interval 1,2,5,10,15
  
  This example will use mail.domain.com as an smtp server, notify users whose password expires in less than 21 days, send mail from support@domain.com
  Report is enabled, reports sent to myaddress@domain.com
  Interval is used, and emails will be sent to people whose password expires in less than 21 days if the script is run, with 15, 10, 5, 2 or 1 days remaining untill password expires.

#>
param(
    # $smtpServer Enter Your SMTP Server Hostname or IP Address
    [Parameter(Mandatory=$True,Position=0)]
    [ValidateNotNull()]
    [string]$smtpServer,
    # Notify Users if Expiry Less than X Days
    [Parameter(Mandatory=$True,Position=1)]
    [ValidateNotNull()]
    [int]$expireInDays,
    [switch]$logging,
    # Log File Path
    [Parameter(Position=4)]
    [string]$logPath,
    # Testing Enabled
    [Parameter(Position=5)]
    [switch]$testing,
    # Test Recipient, eg recipient@domain.com
    [Parameter(Position=6)]
    [string]$testRecipient,
    # Output more detailed status to console
    [Parameter(Position=7)]
    [switch]$status,
    # Log file recipient
    [Parameter(Position=8)]
    [string]$reportto,
    # Notification Interval
    [Parameter(Position=9)]
    [array]$interval
)
###################################################################################################################
# Time / Date Info
$start = [datetime]::Now
$midnight = $start.Date.AddDays(1)
$timeToMidnight = New-TimeSpan -Start $start -end $midnight.Date
$midnight2 = $start.Date.AddDays(2)
$timeToMidnight2 = New-TimeSpan -Start $start -end $midnight2.Date
# System Settings
$textEncoding = [System.Text.Encoding]::UTF8
$today = $start
$Disabled_date = Get-Date -Format "dd/MM/yyyy"

# Insert desired From and mail address here
$from = "Password Notification <noreply@domain.com>"

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
Write-Output "*".PadRight(25,"*")
# Get Users From AD who are Enabled, Passwords Expire and are Not Currently Expired
# To target a specific OU - use the -searchBase Parameter -https://docs.microsoft.com/en-us/powershell/module/addsadministration/get-aduser
# You can target specific group members using Get-AdGroupMember.
# based on earlier version but method still works here.

# Insert DN path here
$OU = 'OU=Users,DC=AD,DC=LOCAL'


$users = $OU | ForEach-Object {get-aduser -searchbase $PSItem -filter {(Enabled -eq $true) -and (PasswordNeverExpires -eq $false)} -properties Name, PasswordNeverExpires, PasswordExpired, PasswordLastSet, EmailAddress}
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
Write-Output "$notifyCount User with expired password in $expireInDays days."
# Process notifyusers
foreach ($user in $notifyUsers)
{
    # Email Address
    $samAccountName = $user.UserName
    # Insert domainname or remove it
    $O365Login = $samAccountName.ToLower() + "@domain.com"
    $emailAddress = $user.EmailAddress
    # Set Greeting Message
    $name = $user.Name
    $messageDays = $user.UserMessage
    # Subject Setting
    $subject="ICT ServiceDesk - password expires $messageDays"
    # Email Body Set Here, Note You can use HTML, including Images.
    $body ="
    <font face=""verdana"">
    Dear $name,
    <p> Your password expires $messageDays<br><br>
    You can change your password at https://account.activedirectory.windowsazure.com/ChangePassword.aspx . You can login with <b>$O365Login</b> and your current password.<br><br>
    Change your password as soon as possible. Please note, if you do not change your password, your account will be automatically disabled.<br><br><br>
    <p>Kind regards, <br><br><br>
    ICT ServiceDesk <br><br><br><br>
    This is an automatically generated email. You can not reply on it.<br>
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
                Send-Mailmessage -smtpServer $smtpServer -from $from -to $emailaddress -subject $subject -body $body -bodyasHTML -priority High -Encoding $textEncoding -ErrorAction Stop
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
            Send-Mailmessage -smtpServer $smtpServer -from $from -to $emailaddress -subject $subject -body $body -bodyasHTML -priority High -Encoding $textEncoding -ErrorAction Stop
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

# Search for users with expired password and disable the account
$users_expired = $OU | ForEach-Object {get-aduser -searchbase $PSItem -filter {(Enabled -eq $true) -and (PasswordNeverExpires -eq $false)} -properties Name, PasswordNeverExpires, PasswordExpired, PasswordLastSet, EmailAddress, Enabled | Where-Object PasswordExpired}


# Count Users
$usersCount_Expired = ($users_Expired | Measure-Object).Count
Write-Output "Found $usersCount_Expired expired User Objects"
# Collect Users
$colUsers_Expired = @()
# Process Each User for Password Expiry
Write-Output "Process User Objects"
foreach ($user_Expired in $users_Expired)
{
    
    $Name = $user_Expired.Name
    $emailaddress = $user_Expired.emailaddress
    $passwordSetDate = $user_Expired.PasswordLastSet
    $samAccountName = $user_Expired.SamAccountName
    $pwdLastSet = $user_Expired.PasswordLastSet
    $pwdExpired = $user_Expired.PasswordExpired
    $Enabled = 'False'
    
    # send e-mail to IT department to notify them of disabled accounts

    # Subject Setting
    $subject_disabled="ICT ServiceDesk - password expired $Name disabled"
    # Email Body Seting
    $body_disabled ="
    <font face=""verdana"">
    Dear IT Infra,
    <p>$Name / $samAccountName is disabled as of $Disabled_date<br><br>
    <p>Kind regards, <br><br><br>
    ICT ServiceDesk <br>
    </font>"
    # Disable User account in AD
    Get-ADuser -Identity $user_Expired | Disable-ADAccount
    # Set note in description field of user
    Set-ADUser -Identity $user_Expired -Add @{description=" Auto disabled - pw expired "+$Disabled_date}
    # Send email to IT Infra (edit mail address to desired recipient)
    Send-Mailmessage -smtpServer $smtpServer -from $from -to it-department@domain.com -subject $subject_disabled -body $body_disabled -bodyasHTML -priority High -Encoding $textEncoding -ErrorAction Stop
    # Store User information

    
    # Create User Object
    $userObj_Expired = New-Object System.Object
    $userObj_Expired | Add-Member -Type NoteProperty -Name UserName -Value $samAccountName
    $userObj_Expired | Add-Member -Type NoteProperty -Name Name -Value $Name
    $userObj_Expired | Add-Member -Type NoteProperty -Name EmailAddress -Value $emailAddress
    $userObj_Expired | Add-Member -Type NoteProperty -Name PasswordSet -Value $pwdLastSet
    $userObj_Expired | Add-Member -Type NoteProperty -Name Expired -Value $pwdExpired
    $userObj_Expired | Add-Member -Type NoteProperty -Name Enabled -Value $Enabled
   
    # Add userObj to colusers array
    $colUsers_Expired += $userObj_Expired
    }
# Count Users
$colUsersCount_Expired = ($colUsers_Expired | Measure-Object).Count
Write-Output "$colusersCount_Expired Users processed"

if($logging)
{
    # Create Log File
    Write-Output "Creating Log File"
    $day = $today.Day
    $month = $today.Month
    $year = $today.Year
    $date = "$day-$month-$year"
    $logFileName = "PasswordLog.csv"
    if(($logPath.EndsWith("\")))
    {
       $logPath = $logPath -Replace ".$"
    }
    $logFile = $logPath, $logFileName -join "\"
    Write-Output "Log Output: $logfile"
    ($notifiedUsers+$colUsers_Expired) | Export-CSV $logFile
    if($reportTo)
    {
        $reportSubject = "Password Expiry Report"
        $reportBody = "Password Expiry Report Attached"
        try{
            Send-Mailmessage -smtpServer $smtpServer -from $from -to $reportTo -subject $reportSubject -body $reportbody -bodyasHTML -priority High -Encoding $textEncoding -Attachments $logFile -ErrorAction Stop 
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
