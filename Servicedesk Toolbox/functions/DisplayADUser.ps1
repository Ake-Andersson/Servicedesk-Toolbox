function displayADUser {
  param(
  )

  $input = $inputBox.Text
  
  try{
    #Search for user and computers
    $user = Get-ADUser -identity $input -Properties *
    $computers = Get-ADComputer -filter "ManagedBy -eq '$user'" -Properties *

    #get Config
    $parentDir = (Get-Item $PSScriptRoot).parent.FullName
    $userCFG = Get-Content "$parentDir\templates\AD_User_template.txt"

    #Display info for User
    foreach($line in $userCFG){
        if($line -like "*{*}" -and $line.Length -ge 4){
            Write-Host $line
            $attribute = $line.Substring($line.IndexOf('{')+1, $line.Length-$line.IndexOf('{')-2)
            $friendlyName = $line.Substring(0, $line.IndexOf('{'))
            $displayAttribute = $user."$attribute"

            if($displayAttribute -ne $null){
                $outputBox.SelectionFont = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Bold)
                $outputBox.AppendText($friendlyName)
                $outputBox.SelectionFont = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Regular)
                $outputBox.AppendText($displayAttribute.toString() + "`r`n")
            }
        }        
    }

    $outputBox.AppendText("`r`n")

    #Display login issues info
    $today = Get-Date
    if($user.AccountExpirationDate -eq $null){
        $outputBox.SelectionColor = 'Green'
        $outputBox.AppendText("✔ - User does not have expiry date`r`n")
    }elseif(-Not ($user.AccountExpirationDate -lt $today)){
        $outputBox.SelectionColor = 'Green'
        $outputBox.AppendText("✔ - User is not expired`r`n")
    }else{
        $outputBox.SelectionColor = 'Red'
        $outputBox.AppendText("✕ - User is expired`r`n")
    }

    if($user.Enabled){
        $outputBox.SelectionColor = 'Green'
        $outputBox.AppendText("✔ - User is not disabled`r`n")
    }else{
        $outputBox.SelectionColor = 'Red'
        $outputBox.AppendText("✕ - User is disabled`r`n")

        $Button6.Tag = "Enable User"
        $Button6.Enabled = $true
        $Button6.Text = "Enable User"
    }
   
    if(-Not($user.LockedOut)){
        $outputBox.SelectionColor = 'Green'
        $outputBox.AppendText("✔ - User is not locked out`r`n")
    }else{
        $outputBox.SelectionColor = 'Red'
        $outputBox.AppendText("✕ - User is locked out`r`n")

        $Button6.Tag = "Unlock User"
        $Button6.Enabled = $true
        $Button6.Text = "Unlock User"
    }

    #if Computers >0, ping computers and display results
    if($computers.cn.Length -ne 0){
        $outputBox.AppendText("`r`nPinging computers...`r`n")

        $foundActive = $false
        $computers = $computers.cn.split(" ")
        foreach($computer in $computers){
            $ping = Test-Connection -ComputerName $computer -Quiet -Count 1
            if($ping){
                $outputBox.AppendText("✔ - " + $computer + " is active `r`n")
                $computerBox.Items.Add($computer)
                $foundActive = $true
            }else{
                $outputBox.AppendText("✕ - " + $computer + " is NOT active `r`n")
            }
        }

        #If an active computer is found, set Buttons accordingly
        if($foundActive){
            $computerBox.SelectedIndex = 0

            $Button1.Tag = "MSRA"
            $Button1.Enabled = $true
            $Button1.Text = "MSRA"

            $Button2.Tag = "MSTSC"
            $Button2.Enabled = $true
            $Button2.Text = "MSTSC"

            $Button3.Tag = "Force Reboot"
            $Button3.Enabled = $true
            $Button3.Text = "Force Reboot"
        }

    }

    $Button4.Tag = "Change Password"
    $Button4.Enabled = $true
    $Button4.Text = "Change Password"

    $Button5.Tag = "Set Expiration Date"
    $Button5.Enabled = $true
    $Button5.Text = "Set Expiration Date"



    <#
    if($usersComputers.Length -gt 0){
        $outputBox.AppendText("`r`nPinging computers...`r`n")
        foreach($computer in $usersComputers){
            $ping = Test-Connection -ComputerName $computer -Quiet -Count 1
            if($ping){
                $outputBox.AppendText("✔ - " + $computer + " is active `r`n")
            }else{
                $outputBox.AppendText("✕ - " + $computer + " is NOT active `r`n")
            }
        }
        
        
    }
    #>
    

  }catch{
    $outputBox.AppendText("`r`n" + $error + "`r`n")
    $error.clear()
    return
  }






}
