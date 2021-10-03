function determineFunction {
  param(
  )

  #Disable findbutton until function is complete  
  $findButton.Enabled = $false

  #Clear everything related to previous search
  $error.Clear()
  $outputBox.Clear()

  $computerBox.Items.Clear()

  $Button1.Tag = ""
  $Button2.Tag = ""
  $Button3.Tag = ""
  $Button4.Tag = ""
  $Button5.Tag = ""
  $Button6.Tag = ""

  $Button1.Enabled = $false
  $Button2.Enabled = $false
  $Button3.Enabled = $false
  $Button4.Enabled = $false
  $Button5.Enabled = $false
  $Button6.Enabled = $false
    
  $Button1.Text = ""
  $Button2.Text = ""
  $Button3.Text = ""
  $Button4.Text = ""
  $Button5.Text = ""
  $Button6.Text = ""
  
  #Check what user is searching for and call corresponding function
  #Functions imported from .\functions\xxxxx.ps1
  if($objectBox.Text -eq "User"){
    displayADUser
  } elseif($objectBox.Text -eq "Computer"){
    displayADComputer
  } elseif($objectBox.Text -eq "Shared Folder"){
    displaySharedFolder
  }

  #Change buttons according to their tags
  
  
  $findButton.Enabled = $true
}

