#Import other files and functions
. "$PSScriptRoot\gui\ToolboxFrame.ps1"
. "$PSScriptRoot\functions\DetermineFunction.ps1"
. "$PSScriptRoot\functions\DisplayADUser.ps1"
. "$PSScriptRoot\functions\DisplayADComputer.ps1"
. "$PSScriptRoot\functions\DisplaySharedFolder.ps1"
. "$PSScriptRoot\functions\DisplayMailbox.ps1"

Import-Module ActiveDirectory

#Function for Button clicks, checks Tag for function
function buttonClick {
  param(
    $button
  )
  
  $selectedItem = $computerBox.SelectedItem

  switch($button.Tag){
    #Start Microsoft Remote Assistance to selected computer
    "MSRA" {& "c:\windows\system32\msra.exe" /offerRA $selectedItem}

    #Start Remote Desktop to selected computer
    "MSTSC" {& "c:\windows\system32\mstsc.exe" /v:$selectedItem}

    #Force Reboot the selected computer
    "Force Reboot" {
        try{
            Restart-Computer -ComputerName $selectedItem -Force
        }catch{
            [System.Windows.Forms.MessageBox]::Show($error , "Error" , [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
        }
        $button.Enabled = $false
    }

    #Change the password of searched user
    "Change Password" {
        $password = [Microsoft.VisualBasic.Interaction]::InputBox("Please input the desired password.", "Enter new password")
        if($password.Length -gt 0){
            try{
                Set-ADAccountPassword -Identity $inputBox.Text -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $password -Force)
            }catch{
                [System.Windows.Forms.MessageBox]::Show($error , "Error" , [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
            }

            if(!$error){
                [System.Windows.Forms.MessageBox]::Show("Password successfully changed!" , "Done!" , [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
            }

            $error.clear()
        }
    }

    #Set expiration date of searched user
    "Set Expiration Date" {
        $expirationDate = [Microsoft.VisualBasic.Interaction]::InputBox("Please input the desired expiration date. (ex 2021-01-01).", "Enter expiration date")
        if($expirationDate.Length -gt 0){
            try{
                Set-ADAccountExpiration -Identity $inputBox.Text -DateTime "$expirationDate"
            }catch{
                [System.Windows.Forms.MessageBox]::Show($error , "Error" , [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
            }
            if(!$error){
                [System.Windows.Forms.MessageBox]::Show("Expiration date successfully set!" , "Done!" , [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
            }

            $error.clear()
        }
    }

    #Enable the searched user
    "Enable User" {
        try{
            Enable-ADAccount -Identity $inputBox.Text
        }catch{
            [System.Windows.Forms.MessageBox]::Show($error , "Error" , [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
        }
        if(!$error){
            [System.Windows.Forms.MessageBox]::Show("User enabled!" , "Done!" , [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
            $button.Enabled = $false
        }
        $error.clear()
    }

    #Unlock the searched user
    "Unlock User" {
        try{
            Unlock-ADAccount -Identity $inputBox.Text
        }catch{
            [System.Windows.Forms.MessageBox]::Show($error , "Error" , [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
        }
        if(!$error){
            [System.Windows.Forms.MessageBox]::Show("User unlocked!" , "Done!" , [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
            $button.Enabled = $false
        }
        $error.clear()
    }

    #Enable the searched user
    "Enable Computer" {
        try{
            Get-ADComputer -Identity $computerBox.SelectedItem | Enable-ADAccount
        }catch{
            [System.Windows.Forms.MessageBox]::Show($error , "Error" , [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
        }
        if(!$error){
            [System.Windows.Forms.MessageBox]::Show("Computer enabled!" , "Done!" , [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
            $button.Enabled = $false
        }
        $error.clear()
    }

    #Query user in CMD window
    "Query User" {
        $target = $computerBox.SelectedItem
        Start-Process cmd -Argument "/K query user /server:$target"
    }

    #Query system information in CMD window
    "Systeminfo" {
        $target = $computerBox.SelectedItem
        Start-Process cmd -Argument "/K systeminfo /s $target"
    }

    #Add AD Group Member
    "Add AD group member" {
        $grp = $computerBox.SelectedItem
        $grpMember = [Microsoft.VisualBasic.Interaction]::InputBox("Please enter a username to add to $grp", "Enter Username")
        if($grpMember.Length -gt 0){
            try{
                Add-ADGroupMember -Identity $grp -Members $grpMember
            }catch{
                [System.Windows.Forms.MessageBox]::Show($error , "Error" , [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
            }

            if(!$error){
                [System.Windows.Forms.MessageBox]::Show("$grpMember has been added to $grp" , "Done!" , [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
            }

            $error.Clear()
        }
    }

    #Remove AD Group Member
    "Remove AD group member" {
        $grp = $computerBox.SelectedItem
        $grpMember = [Microsoft.VisualBasic.Interaction]::InputBox("Please enter a username to remove from $grp", "Enter Username")
        if($grpMember.Length -gt 0){
            try{
                Remove-ADGroupMember -Identity $grp -Members $grpMember -Confirm:$false
            }catch{
                [System.Windows.Forms.MessageBox]::Show($error , "Error" , [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
            }

            if(!$error){
                [System.Windows.Forms.MessageBox]::Show("$grpMember has been removed from $grp" , "Done!" , [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
            }

            $error.Clear()
        }
    }

    #Add Full Access Permission (imported from .\functions\DisplayMailbox.ps1)
    "Add Full Access Permission" {addFullAccessPermission}

    #Remove Full Access Permission (imported from .\functions\DisplayMailbox.ps1)
    "Remove Full Access Permission" {removeFullAccessPermission}

    #Add an autoreply for a mailbox (imported from .\functions\DisplayMailbox.ps1)
    "Set Autoreply" {setAutoReply}

    #Show or remove autoreply for mailbox (imported from .\functions\DisplayMailbox.ps1)
    "Show or Remove Autoreply" {showAutoReply}

  }
}


$inputBox.Add_KeyDown({
    #To enable Enter-press to Find
    if ($_.KeyCode -eq "Enter") {
        $findButton.PerformClick()
        $outputBox.Select()
    }
})

$findButton.Add_Click({
        #call determineFunction (imported from file DetermineFunction.ps1)
        determineFunction
})

$optionsButton.Add_Click({

})



$Button1.Add_Click({
    buttonClick($Button1)
})

$Button2.Add_Click({
    buttonClick($Button2)
})

$Button3.Add_Click({
    buttonClick($Button3)
})

$Button4.Add_Click({
    buttonClick($Button4)
})

$Button5.Add_Click({
    buttonClick($Button5)
})

$Button6.Add_Click({
    buttonClick($Button6)
})


if($exchangeEnabled){
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $exchangeUri -Authentication Kerberos
    Import-PSSession $session
}

$Form.ShowDialog()


if($exchangeEnabled){
    Remove-PSSession $session
}