Write-Host "Branding here." # Branding.

if (-not (Get-Module ActiveDirectory)) # Import the AD module if it isn't already added.
{ 
Import-Module ActiveDirectory -Force
}


function GetAuth #Gets Credentials for later use.
{
$Global:LoggedUser = whoami.exe # Grab the logged in user.
Write-Host "Please enter your office365 details."
$Global:Creds = Get-Credential
Write-Host "Thanks for that, preparing a list of users, it should appear soon!"
}

function SetOOO
{
$ContentDir = "c:\pw\" # Change this to your liking or if you cannot write to the C drive.
$ContentFile1 = "c:\pw\emails.txt" # Change this to your liking or if you cannot write to the C drive.
ForEach ($x in $objListBox.SelectedItems) 
{ 
$selecteduser = $x
Get-ADUser $selecteduser  -Properties EmailAddress| select EmailAddress | Export-Csv $ContentFile1
(Get-Content $ContentFile1) | ForEach-Object { $_ -replace '"' } > $ContentFile1 # The next few lines do some formatting.
(Get-Content $ContentFile1) | ForEach-Object { $_ -replace 'EmailAddress' } > $ContentFile1
(Get-Content $ContentFile1) | ForEach-Object { $_ -replace '#TYPE Selected.Microsoft.ActiveDirectory.Management.ADUser' } > $ContentFile1
(Get-Content $ContentFile1) | ForEach-Object { $_ -replace '' } > $ContentFile1
(Get-Content $ContentFile1) | ? {$_.trim() -ne "" }  > $ContentFile1
}
$emails = (Get-Content $ContentFile1) # Add formatted users to the users array.
$session = New-PSSession –ConfigurationName Microsoft.Exchange –ConnectionUri https://ps.outlook.com/powershell -Authentication Basic -AllowRedirection:$true -Credential $Global:Creds #Setup session.
Import-PSSession $session -AllowClobber # Import the.. wait for it.. session..
$internal = $InternalTextbox.Text
$externall = $ExternalTextbox.Text
$end = $DateTextbox.Text
$d = [datetime]::ParseExact($end, "dd/MM/yyyy", $null) # Parses user input to a format PowerShell can read and store.
Set-MailboxAutoReplyConfiguration -Identity $emails -AutoReplyState Enabled -EndTime $d -InternalMessage $internal -ExternalMessage $externall # Sets the auto replies.
Get-MailboxAutoReplyConfiguration -Identity $emails # Views and reads back the auto reply state.
write-host Out Of Office set for $emails .
}

Function SetupForm 
{ 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") # Loading required assemblies.
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")  
 
$objForm = New-Object System.Windows.Forms.Form  # Creating the form.
$objForm.Text = "Clugston Out Of Office" 
$objForm.Size = New-Object System.Drawing.Size(650,320)  
$objForm.StartPosition = "CenterScreen" 
$btnReset = New-Object System.Windows.Forms.Button 
$btnReset.Location = New-Object System.Drawing.Size(10,225) 
$btnReset.Size = New-Object System.Drawing.Size(75,35) 
$btnReset.Text = "Set Out Of Office" 
$objForm.Controls.Add($btnReset)  # Adds the reset button.

$InternalTextbox = New-Object System.Windows.Forms.TextBox
$InternalTextbox.Size = New-Object System.Drawing.Size(150,172)
$InternalTextbox.Location = New-Object System.Drawing.Size(275,40)
$InternalTextbox.Multiline = 1
$InternalTextbox.Text = "Type Internal message here."
$objForm.Controls.Add($InternalTextbox)

$ExternalTextbox = New-Object System.Windows.Forms.TextBox
$ExternalTextbox.Size = New-Object System.Drawing.Size(150,172)
$ExternalTextbox.Location = New-Object System.Drawing.Size(450,40)
$ExternalTextbox.Multiline = 1
$ExternalTextbox.Text = "Type External message here."
$objForm.Controls.Add($ExternalTextbox)

$DateTextbox = New-Object System.Windows.Forms.TextBox
$DateTextbox.Size = New-Object System.Drawing.Size(170,220)
$DateTextbox.Location = New-Object System.Drawing.Size(100,230)
$DateTextbox.Multiline = 0
$DateTextbox.Text = "DD/MM/YYYY"
$objForm.Controls.Add($DateTextbox)

$objLabel1 = New-Object System.Windows.Forms.Label  
$objLabel1.Location = New-Object System.Drawing.Size(275,225)  
$objLabel1.Size = New-Object System.Drawing.Size(150,30)  
$objLabel1.Text = "Please input an end date in DD/MM/YYYY format." 
$objForm.Controls.Add($objLabel1)  # Adds a label with instructions.
     

$btnReset.Add_Click({  # On button click, go to the reset function.
write-host "Setting Out Of Office." 
SetOOO
})

$CancelButton = New-Object System.Windows.Forms.Button 
$CancelButton.Location = New-Object System.Drawing.Size(550,225) 
$CancelButton.Size = New-Object System.Drawing.Size(75,35) 
$CancelButton.Text = "Exit" 
$CancelButton.Add_Click({ # On click, exit.
Try
{
Remove-Item $ContentFile -Force -Recurse -ea SilentlyContinue # Clean up.
Remove-Item $ContentDir -Force -Recurse -ea SilentlyContinue # Clean up.
}

Catch 
{   
Write-Warning "$($error[0]) "   
Break   
} 
$objForm.Close()
})

$objForm.Controls.Add($CancelButton)  # Adds Cancel Button.
 
$objLabel = New-Object System.Windows.Forms.Label  
$objLabel.Location = New-Object System.Drawing.Size(10,20)  
$objLabel.Size = New-Object System.Drawing.Size(280,20)  
$objLabel.Text = "Please select a user:" 
$objForm.Controls.Add($objLabel)  # Adds a label with instructions.
  
$objListBox = New-Object System.Windows.Forms.ListBox 
$objListBox.Location = New-Object System.Drawing.Size(10,40)  
$objListBox.Size = New-Object System.Drawing.Size(260,300)  
$objListBox.Height = 180 
$objListBox.SelectionMode = "One" # Only allows one user to be selected at a time.
 
$ContentDir = "c:\pw\" # Change this to your liking or if you cannot write to the C drive.
$ContentFile = "c:\pw\Staff.txt" # Change this to your liking or if you cannot write to the C drive.

if((Test-Path $ContentDir) -eq 0) # Checks if $ContentDir exists.
{

Try
{
New-Item $ContentDir -type directory -ea Stop # If it doesn't, it tries to create it. You may need to change the directory and path if you cannot write to the local Drive.
}

Catch
{
Write-Warning "$($error[0]) "
Break
}

}

Try
{ 
Get-ADUser -Filter * -SearchBase "OU=EXAMPLE,OU=EXAMPLE,OU=EXAMPLE,DC=EXAMPLE,DC=local" | Select sAMAccountName | Export-Csv $ContentFile -ea stop  # Grabs all AD Users int he specified DC / OU's.
(Get-Content $ContentFile) | ForEach-Object { $_ -replace '"' } > $ContentFile # The next few lines do some formatting.
(Get-Content $ContentFile) | ForEach-Object { $_ -replace 'sAMAccountName' } > $ContentFile
(Get-Content $ContentFile) | ForEach-Object { $_ -replace '#TYPE Selected.Microsoft.ActiveDirectory.Management.ADUser' } > $ContentFile
(Get-Content $ContentFile) | ForEach-Object { $_ -replace '' } > $ContentFile
(Get-Content $ContentFile) | ? {$_.trim() -ne "" }  > $ContentFile  
[array]$users = (Get-Content $ContentFile) # Add formatted users to the users array.
}   
  
Catch

{   
Write-Warning "$($error[0]) "   
Break   
}      

$uniqueusers = $users | Select-Object -Unique | Sort-Object  # Remove duplicate users.

ForEach($user in $uniqueusers)
{ 
[void] $objListBox.Items.Add($user) # Adds each user to the list box.
} 
$objForm.Controls.Add($objListBox)  # Adds the actual listbox.
$objForm.Topmost = $True 
$objForm.Add_Shown({$objForm.Activate()}) 
[void] $objForm.ShowDialog() # Shows the form!

}

GetAuth
SetupForm 
Pause
