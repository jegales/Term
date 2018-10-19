#############################################################################################
## 
##  PFM Employee Access Termination Tool v2.1
##  -by Joe Gales, October 2, 2018
##  
##  Utility to automate several steps when terminating user accounts
##
##  Input is just the users AD username
##
##  Output logs to $LogFilePath defined in the code.
##
##
#############################################################################################


<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    PFMEATT2
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#region begin GUI{ 

$PFMEATT2                        = New-Object system.Windows.Forms.Form
$PFMEATT2.ClientSize             = '400,400'
$PFMEATT2.text                   = "PFM Employee Access Term Tool v2.1"
$PFMEATT2.TopMost                = $false

$labelEnterUsernameHere          = New-Object system.Windows.Forms.TextBox
$labelEnterUsernameHere.multiline  = $false
$labelEnterUsernameHere.text     = ""
$labelEnterUsernameHere.width    = 300
$labelEnterUsernameHere.height   = 20
$labelEnterUsernameHere.location  = New-Object System.Drawing.Point(48,62)
$labelEnterUsernameHere.Font     = 'Microsoft Sans Serif,12'

$buttonFindUser                  = New-Object system.Windows.Forms.Button
$buttonFindUser.text             = "Find User"
$buttonFindUser.width            = 300
$buttonFindUser.height           = 30
$buttonFindUser.location         = New-Object System.Drawing.Point(48,95)
$buttonFindUser.Font             = 'Microsoft Sans Serif,10'

$ProgressBar1                    = New-Object system.Windows.Forms.ProgressBar
$ProgressBar1.width              = 230
$ProgressBar1.height             = 16
$ProgressBar1.location           = New-Object System.Drawing.Point(82,364)

$PictureBox1                     = New-Object system.Windows.Forms.PictureBox
$PictureBox1.width               = 301
$PictureBox1.height              = 156
$PictureBox1.location            = New-Object System.Drawing.Point(47,182)
$PictureBox1.imageLocation       = "\\hrb-1\hrb1\OPERATIONS\TerminatedUsers\T2.jpg"
$PictureBox1.SizeMode            = [System.Windows.Forms.PictureBoxSizeMode]::zoom

$viewLogButton                   = New-Object system.Windows.Forms.Button
$viewLogButton.BackColor         = "#4a90e2"
$viewLogButton.text              = "View Log"
$viewLogButton.width             = 100
$viewLogButton.height            = 30
$viewLogButton.location          = New-Object System.Drawing.Point(250,10)
$viewLogButton.Font              = 'Microsoft Sans Serif,10'

$resetButton                     = New-Object system.Windows.Forms.Button
$resetButton.BackColor           = "#4a90e2"
$resetButton.text                = "Reset"
$resetButton.width               = 102
$resetButton.height              = 30
$resetButton.location            = New-Object System.Drawing.Point(48,11)
$resetButton.Font                = 'Microsoft Sans Serif,10'

$labelUserStatus                 = New-Object system.Windows.Forms.Label
$labelUserStatus.text            = ""
$labelUserStatus.AutoSize        = $false
$labelUserStatus.width           = 300
$labelUserStatus.height          = 50
$labelUserStatus.location        = New-Object System.Drawing.Point(48,140)
$labelUserStatus.Font            = 'Microsoft Sans Serif,8'

$buttonTerminateUser             = New-Object system.Windows.Forms.Button
$buttonTerminateUser.text        = "Terminate User"
$buttonTerminateUser.width       = 300
$buttonTerminateUser.height      = 30
$buttonTerminateUser.location    = New-Object System.Drawing.Point(48,95)
$buttonTerminateUser.Font        = 'Microsoft Sans Serif,10'

$PFMEATT2.controls.AddRange(@($labelEnterUsernameHere,$buttonFindUser,$ProgressBar1,$PictureBox1,$viewLogButton,$resetButton,$labelUserStatus,$buttonTerminateUser))

$LogFilePath = "\\hrb-1\hrb1\OPERATIONS\TerminatedUsers\TermUser.txt"

#region gui events {
$viewLogButton.Add_Click({ invoke-item $LogFilePath })
#endregion events }

#endregion GUI }

# Connection details for email, the body gets set inside the loop
# If your SMTP server does not support SSL, remove the -UseSSL parameter from the Send-MailMessage cmdlet
$smtpServer = "webmail.pfm.com"
#$port = 25
# Email header
$sendTo = "helpdesk@pfm.com"
$sendFrom = "helpdesk@pfm.com"
$subject = "Termination task for SecretServer"

$MainForm_Load = {
Import-Module ActiveDirectory
}

$buttonFindUser.Add_Click({
$txtUsername = $labelEnterUsernameHere.Text
$date = Get-Date -format "MMM d yyyy"
#Write-Host "txtUsername after buttonFindUser: " $txtUsername
#$usernameCheck = Get-ADUser $txtUsername.Text
$usernameCheck = Get-ADUser $txtUsername
$progressbar1.Value = + 100
$labelUserStatus.ForeColor = 'Red'
If ($usernameCheck -eq $Null) { $labelUserStatus.Text = "User does not exist in AD" }
   Else {
      $labelUserStatus.ForeColor = 'Green'
      $userFound = Get-ADUser -Identity $txtUsername | Select -Expand Name
      $labelUserStatus.Text = " " | Out-File $LogFilePath -Append
      $labelUserStatus.Text = "--------------------------------------------"  #--- separator
      $labelUserStatus.Text | Out-File $LogFilePath -Append                   #--- write the status to file
      $labelUserStatus.Text = $userFound + " " + "has been found on $date!"   #--- set the status for screen
      $labelUserStatus.Text | Out-File $LogFilePath -Append                   #--- write the status to file
      $buttonFindUser.Visible = $false;
      $buttonTerminateUser.Visible = $true
   }
})


$resetButton.Add_Click({
#TODO: Place custom script here
$buttonTerminateUser.Visible = $false;
$buttonFindUser.Visible = $true;
$labelUserStatus.Text = $null;
$progressbar1 = 0;
$labelEnterUsernameHere.Visible = $true;
$labelEnterUsernameHere.Text = ""
$labelUserStatus.Text = $null
$txtUsername = $null
})


$buttonTerminateUser.Add_Click({
Import-Module ActiveDirectory
$txtUsername = $labelEnterUsernameHere.Text
$progressbar1.Value = 0
$usernameCheck = Get-ADUser $txtUsername
If ($usernameCheck -eq $Null) 
   { 
      $labelUserStatus.Text = $txtUsername.Text + " " + "does not exist in AD" 
   }
Else
   {
      $date = Get-Date -format "MMM d yyyy"
      $title = "Terminated Account"
      $userRun = $env:username + " " + "with Employee Access Termination Tool"
      $runBy = $userRun.ToUpper()
      $desc = "Terminated" + " " + $date + " " + "by:" + ' ' + $runBy
      $SecPasswd = ""
      $progressbar1.Value += 10

      $labelUserStatus.Text = $date + " " + $runBy | Out-File $LogFilePath -Append
      $labelUserStatus.Text = " " | Out-File $LogFilePath -Append
      $labelUserStatus.Text = "Retrieve group memberships"
      $labelUserStatus.Text | Out-File $LogFilePath -Append
      $labelUserStatus.Text = " " | Out-File $LogFilePath -Append
      #
      # 1. Retrieve the user in question:
      $progressbar1.Value += 10
      Start-Sleep -s 1
      $User = Get-ADUser $txtUsername -Properties memberOf
      #
      # 2. Retrieve groups that the user is a member of
      $progressbar1.Value += 10
      $Groups = $User.memberOf | ForEach-Object {
      Get-ADGroup $_
      #
      #-----If group contains "SecretServer" then send email to helpdesk
      if ($_.contains("SecretServer"))
         {

# set Email body
$message = @"
Disable user account in Secret Server
User Name: $txtUsername was a member of $_
"@

            #Send-MailMessage -SmtpServer $smtpServer -Credential $credential -To $sendTo -Bcc $sendBcc -From $sendFrom -Subject $subject -Body $message
            #Write-Host "smtpServer:" $smtpServer "sendto:" $sendTo "sendFrom:" $sendFrom "message:"$message "member of:"$_
            Send-MailMessage -SmtpServer $smtpServer -To $sendTo -From $sendFrom -Subject $subject -Body $message
            $labelUserStatus.Text = "Email sent to Helpdesk for SecretServer account deletion"
            $labelUserStatus.Text | Out-File $LogFilePath -Append
         }
      }
      $labelUserStatus.Text = " " | Out-File $LogFilePath -Append
      $labelUserStatus.Text = "Groups Removed" | Out-File $LogFilePath -Append
      $labelUserStatus.Text = " " | Out-File $LogFilePath -Append
      Start-Sleep -s 1
      $labelUserStatus.Text = $User.memberOf | ForEach-Object { Get-ADGroup $_ } | select -Expand Name | Out-File $LogFilePath -Append
      $labelUserStatus.Text = " " | Out-File $LogFilePath -Append
      #
      # 3. Go through the groups and remove the user
      $progressbar1.Value += 10
      $Groups | ForEach-Object { Remove-ADGroupMember -Identity $_ -Members $User -Confirm:$false }
      $labelUserStatus.Text = "Removing" + " " + $txtUsername.Text + " " + "from all groups"
      Start-Sleep -s 1
      $labelUserStatus.Text | Out-File $LogFilePath -Append
      #
      #-----This is where we remove the ip phone field input
      $progressbar1.Value += 10
      $labelUserStatus.Text = "Clearing IP Phone field"
      Start-Sleep -s 1
      $labelUserStatus.Text | Out-File $LogFilePath -Append
      Get-ADuser -Identity $txtUsername | Set-ADuser -Clear ipPhone
      #
      #-----This is where we set the description to terminated user + date
      $progressbar1.Value += 10
      $labelUserStatus.Text = "Renaming description to" + " " + $desc
      Start-Sleep -s 1
      $labelUserStatus.Text | Out-File $LogFilePath -Append
      Set-ADuser -Identity $txtUsername -Description $desc
      #
      #-----This is where we set the description to terminated user + date
      $progressbar1.Value += 15
      $labelUserStatus.Text = "Renaming title to" + " " + $title
      $labelUserStatus.Text | Out-File $LogFilePath -Append
      Start-Sleep -s 1
      Set-ADuser -Identity $txtUsername -Title $title
      #
      #-----This is where we reset the account password (11 Characters long, randomized)
      $SecPasswd = Get-RandomCharacters -length 5 -characters 'abcdefghiklmnoprstuvwxyz'
      $SecPasswd += Get-RandomCharacters -length 3 -characters 'ABCDEFGHKLMNOPRSTUVWXYZ'
      $SecPasswd += Get-RandomCharacters -length 2 -characters '1234567890'
      $SecPasswd += Get-RandomCharacters -length 1 -characters '!"§$%&/()=?}][{@#*+'
      $SecPasswd = Scramble-String $SecPasswd
      $progressbar1.Value = 100
      $labelUserStatus.Text = "Resetting password to a random password"
      Start-Sleep -s 1
      $labelUserStatus.Text | Out-File $LogFilePath -Append
      Set-ADAccountPassword $txtUsername -reset -newpassword (ConvertTo-SecureString -AsPlainText $SecPasswd -Force)
      #
      #-----This is where we disable the account
      Disable-ADAccount -Identity $txtUsername
      $labelUserStatus.Text = $txtUsername + " " + "is now disabled."
      Start-Sleep -s 1
      $labelUserStatus.Text | Out-File $LogFilePath -Append
      #Play sound
      #$sound = new-Object System.Media.SoundPlayer;
      #$sound.SoundLocation = "C:\UTILS\Scripts\TermUser\t3.wav";
      #$sound.Play();
      $labelUserStatus = $txtUsername + " " + "is Terminated"
      $labelUserStatus.Text | Out-File $LogFilePath -Append
   }
})


function Get-RandomCharacters($length, $characters) {
    $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length }
    $private:ofs=""
    return [String]$characters[$random]
}
 
function Scramble-String([string]$inputString){     
    $characterArray = $inputString.ToCharArray()   
    $scrambledStringArray = $characterArray | Get-Random -Count $characterArray.Length     
    $outputString = -join $scrambledStringArray
    return $outputString 
}
 

[void]$PFMEATT2.ShowDialog()