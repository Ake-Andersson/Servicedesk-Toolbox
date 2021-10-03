function displayADComputer {
  param(
  )


  try{
    #Search for computer
    $computer = Get-ADComputer -identity $inputBox.Text -Properties *

    #search for the managing user
    if($computer.ManagedBy -ne $null){
        $user = Get-ADUser -identity $computer.ManagedBy -Properties *
    }

    #get Config
    $parentDir = (Get-Item $PSScriptRoot).parent.FullName
    $computerCFG = Get-Content "$parentDir\templates\AD_Computer_template.txt"

    #Display info for computer
    foreach($line in $computerCFG){
        $attribute = $line.Substring($line.IndexOf('{')+1, $line.Length-$line.IndexOf('{')-2)
        $friendlyName = $line.Substring(0, $line.IndexOf('{'))

        $outputBox.SelectionFont = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Bold)
        $outputBox.AppendText($friendlyName)
        $outputBox.SelectionFont = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Regular)
        $outputBox.AppendText($computer."$attribute".toString() + "`r`n")
    }

    if($user -ne $null){
        $outputBox.AppendText("`r`n")
        $outputBox.SelectionFont = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Bold)
        $outputBox.AppendText("Managing User: ")
        $outputBox.SelectionFont = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Regular)
        $outputBox.AppendText($user.cn + "`r`n")
    }

    $outputBox.AppendText("`r`n")
    if($computer.Enabled){
        $outputBox.SelectionColor = 'Green'
        $outputBox.AppendText("✔ - Computer is not disabled`r`n")
    }else{
        $outputBox.SelectionColor = 'Red'
        $outputBox.AppendText("✕ - Computer is disabled`r`n")

        $Button6.Tag = "Enable Computer"
        $Button6.Enabled = $true
        $Button6.Text = "Enable Computer"
    }

    $outputBox.AppendText("`r`n")
    #Ping computer, and if active set buttons
    $ping = Test-Connection -ComputerName $computer.cn -Quiet -Count 1
    if($ping){
        $outputBox.AppendText("✔ - " + $computer.cn + " is active!`r`n`r`n")

        $IP = [System.Net.Dns]::GetHostAddresses($computer.cn)
        $outputBox.SelectionFont = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Bold)
        $outputBox.AppendText("IP Address: ")
        $outputBox.SelectionFont = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Regular)
        $outputBox.AppendText($IP + "`r`n")

        $computerBox.Items.Add($computer.cn)
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

        $Button4.Tag = "Query User"
        $Button4.Enabled = $true
        $Button4.Text = "Query User (CMD)"

        $Button5.Tag = "Systeminfo"
        $Button5.Enabled = $true
        $Button5.Text = "Systeminfo (CMD)"
    }else{
        $outputBox.AppendText("✕ - " + $computer.cn + " is NOT active!")
    }
    

  }catch{
    $outputBox.AppendText("`r`n" + $error + "`r`n")
    $error.clear()
    return
  }



}