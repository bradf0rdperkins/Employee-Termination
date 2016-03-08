  # Load the Winforms assembly
  [reflection.assembly]::LoadWithPartialName( "System.Windows.Forms")
 
  $global:currenttime= Set-PSBreakpoint -Variable currenttime -Mode Read -Action { $global:currenttime= Get-Date }
  $global:my_username=""
  $my_message = ""
  $global:keepgoing=$true
  $validusername=$true
 
  # CREATE FORM ####################################################
  $form = New-Object Windows.Forms.Form
  $Form.AutoSize = $True
  $Form.AutoSizeMode = "GrowAndShrink"
  $form.text = "Employee Terminator"
 
  $userlabel = New-Object Windows.Forms.Label
  $userlabel.Location = New-Object Drawing.Point 10,10
  $userlabel.Size = New-Object Drawing.Point 50,28
  $userlabel.text = "User:"
  $usertextfield = New-Object Windows.Forms.TextBox
  $usertextfield.Location = New-Object Drawing.Point 65,10
  $usertextfield.Size = New-Object Drawing.Point 150,15

  $managerlabel = New-Object Windows.Forms.Label
  $managerlabel.Location = New-Object Drawing.Point 10,40
  $managerlabel.Size = New-Object Drawing.Point 52,18
  $managerlabel.text = "Manager:"
  $managertextfield = New-Object Windows.Forms.TextBox
  $managertextfield.Location = New-Object Drawing.Point 65,40
  $managertextfield.Size = New-Object Drawing.Point 150,15
 
  $outputtextbox = New-Object Windows.Forms.TextBox
  $outputtextbox.Location = New-Object Drawing.Point 10,70
  $outputtextbox.Size = New-Object Drawing.Point 350,200
  $outputtextbox.multiline=$true
  $outputtextbox.scrollbars="Vertical"
 
  $closebutton = New-Object Windows.Forms.Button
  $closebutton.text = "Close"
  $closebutton.Location = New-Object Drawing.Point 375,35
 
  $gobutton = New-Object Windows.Forms.Button
  $gobutton.text = "Disable"
  $gobutton.Location = New-Object Drawing.Point 375,8
 
  $form.controls.add($userlabel)
  $form.controls.add($usertextfield)
  $form.controls.add($managerlabel)
  $form.controls.add($managertextfield)
  $form.controls.add($gobutton)
  $form.controls.add($closebutton)
  $form.controls.add($outputtextbox)
  $form.AcceptButton = $gobutton
 
  ########################################################################################################
 
  # FORM FUNTIONALITY ####################################################################################
  import-module ActiveDirectory
  Add-PSSnapin Microsoft.Exchange.Management.PowerShell.Admin

  function OutToBox ($OutText)
  {
    $outputtextbox.appendtext($outtext)
    [System.Windows.Forms.Application]::DoEvents()
  }
 

  # close BUTTON ##
  $closebutton.add_click({
        [System.Windows.Forms.Application]::Exit($null)
  })

 
  # Disable Button ##
  $gobutton.add_click({
   
    $closebutton.enabled=$true
 
    $validusername=$true
    $validmanager=$true
    $my_username=$usertextfield.Text
    $manager=$managertextfield.Text  
    $global:keepgoing=$true
    $Date = (Get-Date -format MM/dd/yyyy)

    #Set variables for email
    $mailbox = get-mailbox -identity $my_username
    $Supervisor = $manager
    $SupEmail = $Supervisor + "@company.com"
    #$SupFullName = Get-ADUser -Filter {name -like "$($Supervisor.DisplayName)"} -Properties * | select samaccountname
    $UPN = "@companydomain.com"
    $NTDomain = "companydomain"
    $ExchangeServer = "mailserver"
    $EmailDomain = "@companydomain.com"
    $EmailAddress = "user@company.com"
    $From = "Administrators@company.com"
    $SMTPServer = "mail.company.com"
    $Body = $my_username + "'s account has been disabled and you have been granted permissions to their mailbox. Please contact IT with questions or submit a helpdesk ticket for assistance with accessing the mailbox.`r`n`r`nThank you,`r`ncompany IT Administration Team"
   
    #Evaluate if the user exists
    try{Get-ADUser $my_username}
    catch
    {
        $validusername=$false
    }
     
    if ($validusername -ne 0)
    {
        $gobutton.Enabled = $false
        $outputtextbox.appendtext("Working...`r`n")
        
        #Disable OWA and Activesync to prevent access to email
        Set-Mailbox -Identity $my_username -RecipientLimits 0
         $my_message = "Recipient limit set to 0`r`n"
         OutToBox($my_message)
        Set-CASMailbox -Identity $my_username -OWAEnabled:$False
         $my_message = "OWA disabled`r`n"
         OutToBox($my_message)
        Set-CASMailbox -Identity $my_username -ActiveSyncEnabled:$False
         $my_message = "ActiveSync disabled`r`n"
         OutToBox($my_message)
 
        #Disable AD account
        Disable-ADAccount -Identity $my_username
         $my_message = "Active Directory account disabled`r`n"
         OutToBox($my_message)
        
        #Change description to date disabled
        $Regex='^(0[1-9]|1[012])[- /.](0[1-9]|[12][0-9]|3[01])[- /.](20)\d\d$'
        If ($my_username.Description -notmatch $Regex)
        {
            Set-ADUser -Identity $my_username -Description $Date
             $my_message = "Set Description to today's date: " + $Date + " `r`n"
             OutToBox($my_message)
        }
 
        #Move user to the disabled accounts OU
        Get-ADUser $my_username | Move-ADObject -TargetPath "OU=Users,OU=Disabled Accounts,DC=companydomain,DC=com"
         $my_message = "Moved " + $my_username + " to the Disabled Accounts\Users OU `r`n"
         OutToBox($my_message)

        # Remove user from groups 
        $groups = get-adprincipalgroupmembership $my_username
        # Loop through each group
        foreach ($group in $groups) {
            # Exclude Domain Users group
            if ($group.name -ne "domain users") {
                # Remove user from group
                remove-adgroupmember -Identity $group.name -Member $my_username -Confirm:$false
                # Write progress to screen
                $my_message = "Removed " + $my_username + " from " + $group.name + "`r`n"
                OutToBox($my_message)
                # Define and save group names into filename
                $grouplogfile = "\\server\Private\AD Scripts\Logs\" + $my_username + "_removedGroups.txt"
                $group.name >> $grouplogfile
            }
        }

        #Check if manager exists
        try{Get-ADUser $manager}
            catch
            {
                $validmanager=$false
            }
            if ($validmanager -ne 0)
            {
                #Grant permissions to the user's supervisor
                $Mailbox | Add-MailboxPermission -user $Supervisor -AccessRights FullAccess
                $my_message = $Supervisor + " now has access to " + $my_username + "'s mailbox`r`n"
                OutToBox($my_message)
    
                $gobutton.Enabled = $true
                }
            else
            {
                if ($validmanager -ne $true)
                {
                    outtobox("Manager does not exist`r`n")
                }
            }

        #Send email to supervisor notifying them of the disabled account and granted permissions
        Send-MailMessage -To $SupEmail -From $From -Subject "Account Disabled [$my_username]" -Body $Body -SmtpServer $SMTPServer
         $my_message = "Sending email to " + $Supervisor + "`r`n"
         OutToBox($my_message)

        $outputtextbox.appendtext("Done!`r`n")
        $gobutton.Enabled = $true
    }
    else
    {
        if ($validusername -ne $true)
        {
            outtobox("User does not exist`r`n")
        }
    }
  })
 
  ###############################################################
 
  # DISPLAY DIALOG
  $form.ShowDialog()
