#Goal of the function is to fetch all users with access rights to a folder
#Check if each access right is an AD group, and if so, display members
function displaySharedFolder {
  param(
  )
  
    $path = $inputBox.Text
    $directUsers = @()
    
    #Try to fetch access rights to group
    try{
        $accessRights = (Get-ACL $path).access | select IdentityReference,FileSystemRights,AccessControlType,IsInherited,InheritanceFlags
    }catch{
        $outputBox.AppendText($error)
        $error.Clear()
        return
    }

    if($accessRights -eq $null){
        $outputBox.AppendText("An error occured! Incorrect path?")
        return
    }

    $outputBox.AppendText("Folder: " + $inputBox.Text + "`r`n")

    #Display access righs, and try to fetch AD group for each right
    $foundGroup = $false
    $outputBox.SelectionFont = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Bold)
    $outputBox.AppendText("`r`nThe following groups have access to the folder:`r`n")

    foreach($access in $accessRights){
        $outputBox.SelectionFont = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Regular)

        try{
            $testName = $access.IdentityReference.ToString()
            $testName = $testName.Substring(($testName.IndexOf('\')+1), $testName.Length-$testName.IndexOf('\')-1) #remove "domain\"

            $group = Get-ADGroup -Identity $testName
        }catch{
            #accessright is NOT an AD group
            $directUsers += $access
        }

        if(!$error){
            #accessright IS an AD group, so fetch members
            if($group.name -ne "Administrators"){
                $computerBox.Items.Add($group.name)
                $foundGroup = $true
            }

            $outputBox.AppendText("`r`n")
            $outputBox.SelectionFont = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Bold)
            $outputBox.AppendText("(" + $access.FileSystemRights.toString() + ") ")
            if($access.isInherited.toString() -eq "True"){
                $outputBox.AppendText("[Inherited]) - " + $access.IdentityReference.toString() + ":`r`n")
            }else{
                $outputBox.AppendText("[Not Inherited] - " + $access.IdentityReference.toString() + ":`r`n")
            }

            try{
                $members = Get-ADGroupMember -Identity $group -Recursive | Get-ADUser | select sAMAccountName
            }catch{
                $outputBox.AppendText("Failed getting members of " + $group + ":`r`n" + $error + "`r`n")
            }

            #display members
            if(!$error){
                foreach($member in $members){
                    $outputBox.SelectionFont = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Regular)
                    $outputBox.AppendText($member.sAMAccountName + "`r`n")
                }
            }

        }

        $error.Clear()
    }

    #Display access rights that are not groups
    if($directUsers.Length -gt 0){
        $outputBox.SelectionFont = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Bold)
        $outputBox.AppendText("`r`nThe following users have access to the folder directly:`r`n")

        foreach($user in $directUsers){
            $outputBox.SelectionFont = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Regular)
            $outputBox.AppendText("(" + $user.FileSystemRights.toString() + ") ")
            if($user.isInherited.toString() -eq "True"){
                $outputBox.AppendText("[Inherited]) - " + $user.IdentityReference.toString() + "`r`n")
            }else{
                $outputBox.AppendText("[Not Inherited] - " + $user.IdentityReference.toString() + "`r`n")
            }
        }
    }

    #Set button functions
    if($foundGroup){
        $computerBox.SelectedIndex = 0

        $Button1.Tag = "Add AD group member"
        $Button1.Enabled = $true
        $Button1.Text = "Add group member"

        $Button2.Tag = "Remove AD group member"
        $Button2.Enabled = $true
        $Button2.Text = "Remove group member"
    }

}