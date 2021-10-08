function displayMailbox {
  param(
  )

  $outputBox.Text = "Loading..."
  $input = $inputBox.Text

  try{
    $mailbox = Get-Mailbox -identity $input
    $folderStatistics = Get-MailboxFolderStatistics -identity $mailbox.SamAccountName | select Name, FolderandSubFolderSize,ItemsinFolderandSubfolders
    $userPermissions =  (Get-MailboxPermission -identity $mailbox.SamAccountName | where { ($_.AccessRights -eq "FullAccess") -and ($_.IsInherited -eq $false) -and -not ($_.User -like "NT AUTHORITY\SELF") } )
    $mailboxStatistics = Get-MailboxStatistics -Identity $mailbox.SamAccountName | select TotalItemSize
  }catch{
    $outputBox.AppendText($error)
    $error.clear()  
    return
  }

  if($mailbox -eq $null){
    $outputBox.AppendText("Error! Mailbox not found?")
    return
  }

  $outputBox.Text = ""

  #Display Mailbox info
  $outputBox.SelectionFont = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Bold)
  $outputBox.AppendText("Username: ")
  $outputBox.SelectionFont = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Regular)
  $outputBox.AppendText($mailbox.sAmAccountName + "`r`n")

  $outputBox.SelectionFont = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Bold)
  $outputBox.AppendText("Database: ")
  $outputBox.SelectionFont = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Regular)
  $outputBox.AppendText($mailbox.database + "`r`n")

  if($mailbox.emailAddresses -ne $null){
    $emailAddresses = $mailbox.emailAddresses -replace "smtp:", ""
    $outputBox.SelectionFont = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Bold)
    $outputBox.AppendText("Email Adresses: `r`n")
    foreach($address in $emailAddresses){
        $outputBox.AppendText($address + "`r`n")
    }
  }

  #Display Mailbox size
  if($mailboxStatistics -ne $null){
    $outputBox.AppendText("`r`n")
    $outputBox.SelectionFont = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Bold)
    $outputBox.AppendText("Total size: ")
    $outputBox.SelectionFont = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Regular)
    $outputBox.AppendText($mailboxStatistics.TotalItemSize.toString() + "`r`n")
  }

  $outputBox.AppendText("`r`n")

  #Display User Permissions
  if($userPermissions -ne $null){
    $outputBox.SelectionFont = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Bold)
    $outputBox.AppendText("Non-inherited Full access permissions: `r`n")

    foreach($permission in $userPermissions){
        $outputBox.AppendText($permission.User + "`r`n")
    }

    $outputBox.AppendText("`r`n")
  }

  #Display mailbox folders and sizes
  if($folderStatistics -ne $null){
    $outputBox.SelectionFont = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Bold)
    $outputBox.AppendText("Folders: " + $folderStatistics.Length + "`r`n")

    foreach($folder in $folderStatistics){
        $outputBox.AppendText($folder.name + ": " + $folder.ItemsinFolderandSubfolders + " items - " + $folder.FolderandSubFolderSize + "`r`n")
    }
  }

  $computerBox.Items.Add($mailbox.SamAccountName)
  $computerBox.SelectedIndex = 0

  #Set buttons
  $Button1.Tag = "Add Full Access Permission"
  $Button1.Enabled = $true
  $Button1.Text = "Add Full Access Permission"

  $Button2.Tag = "Remove Full Access Permission"
  $Button2.Enabled = $true
  $Button2.Text = "Remove Full Access Permission"

  $Button4.Tag = "Set Autoreply"
  $Button4.Enabled = $true
  $Button4.Text = "Set Autoreply"

  $Button5.Tag = "Show or Remove Autoreply"
  $Button5.Enabled = $true
  $Button5.Text = "Show or Remove Autoreply"

}

#Function to add a user to mailbox full access permissions
function addFullAccessPermission{
    $userName = [Microsoft.VisualBasic.Interaction]::InputBox("Please enter the username of the user you wish to add.", "Enter username")

    if($userName.Length -gt 0){
        try{
            Add-MailboxPermission -identity $computerBox.SelectedItem -accessrights:fullaccess -user $userName
        }catch{
            [System.Windows.Forms.MessageBox]::Show($error , "Error" , [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
            $error.Clear()
            return
        }

        if(!$error){
            [System.Windows.Forms.MessageBox]::Show("$userName has been added to the mailbox Full Access Permissions" , "Done!" , [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
        }
    }

    $error.clear()
}


#Function to remove a user to mailbox full access permissions
function removeFullAccessPermission{
    $userName = [Microsoft.VisualBasic.Interaction]::InputBox("Please enter the username of the user you wish to add.", "Enter username")
    
    if($userName.Length -gt  0){
        try{
            Remove-MailboxPermission -identity $computerBox.SelectedItem -accessrights:fullaccess -user $userName -confirm:$false
        }catch{
            [System.Windows.Forms.MessageBox]::Show($error , "Error" , [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
            $error.Clear()
            return
        }

        if(!$error){
            [System.Windows.Forms.MessageBox]::Show("$userName has been removed to the mailbox Full Access Permissions" , "Done!" , [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
        }

        $error.clear()
    }

}


#Function to set autoreply for mailbox
function setAutoReply{
    $replyText = ""

    $autoreplyForm                   = New-Object system.Windows.Forms.Form
    $autoreplyForm.ClientSize        = '400,300'
    $autoreplyForm.text              = "Autoreply"
    $autoreplyForm.FormBorderStyle   = 'FixedDialog'

    $autoreplyTextBox                = New-Object system.Windows.Forms.TextBox
    $autoreplyTextBox.multiline      = $true
    $autoreplyTextBox.width          = 380
    $autoreplyTextBox.height         = 223
    $autoreplyTextBox.location       = New-Object System.Drawing.Point(10,10)
    $autoreplyTextBox.Font           = 'Microsoft Sans Serif,10'

    $datePicker                      = New-Object Windows.Forms.DateTimePicker
    $datePicker.ShowUpDown           = $false
    $datePicker.width                = 80
    $datePicker.height               = 30
    $datePicker.location             = New-Object System.Drawing.Point(10,276)
    $datePicker.Enabled              = $true

    $startLabel                      = New-Object system.Windows.Forms.Label
    $startLabel.text                 = "Startdate"
    $startLabel.width                = 80
    $startLabel.height               = 30
    $startLabel.location             = New-Object System.Drawing.Point(7,257)
    $startLabel.Font                 = 'Microsoft Sans Serif,10'

    $datePicker2                     = New-Object Windows.Forms.DateTimePicker
    $datePicker2.ShowUpDown          = $false
    $datePicker2.width               = 80
    $datePicker2.height              = 30
    $datePicker2.location            = New-Object System.Drawing.Point(100,276)
    $datePicker2.Enabled             = $true

    $scheduledBox                    = New-Object system.Windows.Forms.CheckBox
    $scheduledBox.text               = "Schedule with start/end-dates"
    $scheduledBox.AutoSize           = $true
    $scheduledBox.width              = 95
    $scheduledBox.height             = 20
    $scheduledBox.location           = New-Object System.Drawing.Point(7,238)
    $scheduledBox.Font               = 'Microsoft Sans Serif,8'
    $scheduledBox.Checked            = $true
    $scheduledBox.Add_Click({
        $datePicker.Enabled = !$datePicker.Enabled
        $datePicker2.Enabled = !$datePicker2.Enabled
    })

    
    $internalBox                    = New-Object system.Windows.Forms.CheckBox
    $internalBox.text               = "Only internal autoreply"
    $internalBox.AutoSize           = $true
    $internalBox.width              = 95
    $internalBox.height             = 20
    $internalBox.location           = New-Object System.Drawing.Point(200,238)
    $internalBox.Font               = 'Microsoft Sans Serif,8'
    $internalBox.Add_Click({

    })

    $endLabel                        = New-Object system.Windows.Forms.Label
    $endLabel.text                   = "Enddate"
    $endLabel.width                  = 80
    $endLabel.height                 = 30
    $endLabel.location               = New-Object System.Drawing.Point(97,257)
    $endLabel.Font                   = 'Microsoft Sans Serif,8'

    $okButton                        = New-Object system.Windows.Forms.Button
    $okButton.text                   = "OK"
    $okButton.width                  = 80
    $okButton.height                 = 30
    $okButton.location               = New-Object System.Drawing.Point(211,261)
    $okButton.Font                   = 'Microsoft Sans Serif,8'
    $okButton.Add_Click({
        $replyText = $autoreplyTextBox.Text
        if(-not($replyText -eq $null)){
            $mail = $inputBox.Text
            $replyText = $replyText.replace("`n","<br>")
            $date1 = $datePicker.Value
            $date2 = $datePicker2.Value

            try{
                if($scheduledBox.Checked){
                    if($internalBox.Checked){
                        #Internal message only, with schedule 
                        Set-MailboxAutoReplyConfiguration -Identity $computerBox.SelectedItem -AutoReplyState Scheduled -StartTime $date1 -EndTime $date2 -InternalMessage $replyText
                    }else{
                        #External and internal message, with schedule
                        Set-MailboxAutoReplyConfiguration -Identity $computerBox.SelectedItem -AutoReplyState Scheduled -StartTime $date1 -EndTime $date2 -InternalMessage $replyText -ExternalMessage $replyText
                    }
                }else{
                    if($internalBox.Checked){
                        #Internal message only, without schedule
                        Set-MailboxAutoReplyConfiguration -Identity $computerBox.SelectedItem -AutoReplyState Enabled -InternalMessage $replyText
                    }else{
                        #External and internal message, without schedule
                        Set-MailboxAutoReplyConfiguration -Identity $computerBox.SelectedItem -AutoReplyState Enabled -InternalMessage $replyText -ExternalMessage $replyText
                    }
                }
            }catch{
                [System.Windows.Forms.MessageBox]::Show($error , "Error" , [System.Windows.Forms.MessageBoxButtons]::OK)
            }

            if(!$error){
                [System.Windows.Forms.MessageBox]::Show("Autoreply has now been set!" , "Done!" , [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
            }

            $error.clear()
            $autoReplyForm.Close()
            return
        }else{
            [System.Windows.Forms.MessageBox]::Show("No reply has been entered!" , "Error" , [System.Windows.Forms.MessageBoxButtons]::OK)
        }
    })

    $cancelButton                    = New-Object system.Windows.Forms.Button
    $cancelButton.text               = "Cancel"
    $cancelButton.width              = 80
    $cancelButton.height             = 30
    $cancelButton.location           = New-Object System.Drawing.Point(310,261)
    $cancelButton.Font               = 'Microsoft Sans Serif,10'
    $cancelButton.Add_Click({
        $autoreplyForm.Close()
        return
    })

    $autoreplyForm.controls.AddRange(@($okButton,$cancelButton,$autoreplyTextBox,$datePicker, $startLabel, $datePicker2, $endLabel, $scheduledBox, $internalBox))

    [void] $autoreplyForm.ShowDialog()
}


#Function to show autoreply for mailbox
function showAutoReply{
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $replyText = Get-MailboxAutoReplyConfiguration -Identity $computerBox.SelectedItem | select StartTime,EndTime,ExternalMessage,InternalMessage,AutoReplyState
    
    if($replyText -eq $null){
        [System.Windows.Forms.MessageBox]::Show("There is no autoreply set for this mailbox!" , "Error" , [System.Windows.Forms.MessageBoxButtons]::OK)
        return
    }

    $autoreplyForm                   = New-Object system.Windows.Forms.Form
    $autoreplyForm.ClientSize        = '400,300'
    $autoreplyForm.text              = "Current autoreply"
    $autoreplyForm.FormBorderStyle   = 'FixedDialog'

    $autoreplyTextBox                = New-Object system.Windows.Forms.TextBox
    $autoreplyTextBox.multiline      = $true
    $autoreplyTextBox.width          = 380
    $autoreplyTextBox.height         = 223
    $autoreplyTextBox.location       = New-Object System.Drawing.Point(10,10)
    $autoreplyTextBox.Font           = 'Microsoft Sans Serif,10'
    $autoreplyTextBox.ReadOnly       = $true

    $internalBox                    = New-Object system.Windows.Forms.CheckBox
    $internalBox.text               = "Only internal autoreply"
    $internalBox.AutoSize           = $true
    $internalBox.width              = 95
    $internalBox.height             = 20
    $internalBox.location           = New-Object System.Drawing.Point(200,238)
    $internalBox.Font               = 'Microsoft Sans Serif,8'
    $internalBox.Enabled            = $false

    $scheduledBox                    = New-Object system.Windows.Forms.CheckBox
    $scheduledBox.text               = "Schedule with start/end-dates"
    $scheduledBox.AutoSize           = $true
    $scheduledBox.width              = 95
    $scheduledBox.height             = 20
    $scheduledBox.location           = New-Object System.Drawing.Point(7,238)
    $scheduledBox.Font               = 'Microsoft Sans Serif,8'
    $scheduledBox.Checked            = $true
    $scheduledBox.Enabled            = $false

    $startLabel                      = New-Object system.Windows.Forms.Label
    $startLabel.text                 = "Startdate"
    $startLabel.width                = 80
    $startLabel.height               = 30
    $startLabel.location             = New-Object System.Drawing.Point(7,257)
    $startLabel.Font                 = 'Microsoft Sans Serif,10'

    $endLabel                        = New-Object system.Windows.Forms.Label
    $endLabel.text                   = "Enddate"
    $endLabel.width                  = 80
    $endLabel.height                 = 30
    $endLabel.location               = New-Object System.Drawing.Point(97,257)
    $endLabel.Font                   = 'Microsoft Sans Serif,10'

    if(-not (($replyText.ExternalMessage).Length -eq 0)){
        $text = $replyText.ExternalMessage
        $text = $text -Replace("<br>","`r`n")
        $text = $text -Replace("<html>","")
        $text = $text -Replace("<body>","")
        $text = $text -Replace("</html>","")
        $text = $text -Replace("</body>","")
        $autoreplyTextBox.Text = $text        
    }elseif(-not (($replyText.InternalMessage).Length -eq 0)){
        $text = $replyText.InternalMessage
        $text = $text -Replace("<br>","`r`n")
        $text = $text -Replace("<html>","")
        $text = $text -Replace("<body>","")
        $text = $text -Replace("</html>","")
        $text = $text -Replace("</body>","")
        $internalBox.Checked = $true
        $autoreplyTextBox.Text = $text + "`r`n`r`n[THIS IS AN INTERNAL AUTOREPLY]"    
    }

    if($replyText.AutoReplyState -eq "Enabled"){
        $scheduledBox.Checked = $false
        $startLabel.Visible = $false
        $endLabel.Visible = $false
    }elseif($replyText.AutoReplyState -eq "Scheduled"){
        $datePicker                      = New-Object Windows.Forms.DateTimePicker
        $datePicker.ShowUpDown           = $false
        $datePicker.width                = 80
        $datePicker.height               = 30
        $datePicker.location             = New-Object System.Drawing.Point(10,276)
        $datePicker.Enabled              = $false
        $datePicker.Format = [windows.forms.datetimepickerFormat]::custom
        $datePicker.CustomFormat = “yyyy-MM-dd”
        $datePicker.Text = $replyText.StartTime

        $datePicker2                     = New-Object Windows.Forms.DateTimePicker
        $datePicker2.ShowUpDown          = $false
        $datePicker2.width               = 80
        $datePicker2.height              = 30
        $datePicker2.location            = New-Object System.Drawing.Point(100,276)
        $datePicker2.Enabled             = $false
        $datePicker2.Format = [windows.forms.datetimepickerFormat]::custom
        $datePicker2.CustomFormat = “yyyy-MM-dd”
        $datePicker2.Text = $replyText.EndTime
    }else{
        [System.Windows.Forms.MessageBox]::Show("There is no autoreply set for this mailbox!" , "Error" , [System.Windows.Forms.MessageBoxButtons]::OK)
        $autoreplyForm.Close()
        return
    }

    $okButton                        = New-Object system.Windows.Forms.Button
    $okButton.text                   = "Disable"
    $okButton.width                  = 80
    $okButton.height                 = 30
    $okButton.location               = New-Object System.Drawing.Point(211,261)
    $okButton.Font                   = 'Microsoft Sans Serif,10'
    $okButton.Add_Click({
        try{
            Set-MailboxAutoReplyConfiguration $inputBox.Text –AutoReplyState Disabled –ExternalMessage $null –InternalMessage $null
        }catch{
            [System.Windows.Forms.MessageBox]::Show($error , "Error" , [System.Windows.Forms.MessageBoxButtons]::OK)
        }

        if(!$error){
            [System.Windows.Forms.MessageBox]::Show("Autoreply has now been removed/disabled!" , "Done!" , [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
        }

        $error.clear()
        $autoreplyForm.Close()
        return
    })

    $cancelButton                    = New-Object system.Windows.Forms.Button
    $cancelButton.text               = "Cancel"
    $cancelButton.width              = 80
    $cancelButton.height             = 30
    $cancelButton.location           = New-Object System.Drawing.Point(310,261)
    $cancelButton.Font               = 'Microsoft Sans Serif,10'
    $cancelButton.Add_Click({
        $autoreplyForm.Close()
        return
    })

    $autoreplyForm.controls.AddRange(@($okButton,$cancelButton,$autoreplyTextBox,$datePicker, $startLabel, $datePicker2, $endLabel, $internalBox, $scheduledBox))

    [void] $autoreplyForm.ShowDialog()
}
