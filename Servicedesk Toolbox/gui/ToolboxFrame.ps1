[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#-------------------------------------------Enable Exchange-------------------------------------------
#Check if exchange is enabled in config, and save the uri to variable
$parentDir = (Get-Item $PSScriptRoot).parent.FullName
$exchangeEnabled = $false
$exchangeUri = ""
$exchangeConfig =  Get-Content "$parentDir\configs\Exchange_config.txt"

foreach($line in $exchangeConfig){
    if($line -like "*Enable Exchange*" -and $line -like "*True*"){
        $exchangeEnabled = $true
    }elseif($line -like "*Exchange Powershell-Uri*"){
        $exchangeUri = $line.substring($line.indexOf('=')+1, $line.Length - $line.indexOf('=')-1)
    }
}

Write-Host Exchange enabled: $exchangeEnabled
Write-Host Exchange uri: $exchangeUri
#-------------------------------------------Enable Exchange END---------------------------------------

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '550,510'
$Form.FormBorderStyle            = 'FixedDialog'
$background                      = [system.drawing.image]::FromFile("$PSScriptRoot\img\background.png") #Background image for the frame
$Form.BackgroundImage            = $background
$Form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon("$PSScriptRoot\img\toolbox.ico")
$Form.text                       = "Servicedesk Toolbox"

$inputBox                        = New-Object system.Windows.Forms.TextBox
$inputBox.multiline              = $false
$inputBox.width                  = 160
$inputBox.height                 = 20
$inputBox.location               = New-Object System.Drawing.Point(115,15)
$inputBox.Font                   = 'Microsoft Sans Serif,10'

$objectBox                       = New-Object System.Windows.Forms.ComboBox
$objectBox.width                 = 100
$objectBox.height                = 20
$objectBox.location              = New-Object System.Drawing.Point(10,15)
$objectBox.Font                  = New-Object System.Drawing.Font("Arial",9,[System.Drawing.FontStyle]::Regular)
$objectBox.DropDownStyle         = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$objectOptions = @("User","Computer","Shared Folder")
foreach($opt in $objectOptions){
    $objectBox.Items.Add($opt)
}
if($exchangeEnabled){
    $objectBox.Items.Add("Mailbox")
}
$objectBox.SelectedIndex         = 0

$findButton                      = New-Object system.Windows.Forms.Button
$findButton.text                 = "Search!"
$findButton.width                = 110
$findButton.height               = 40
$findButton.location             = New-Object System.Drawing.Point(280,5)
$findButton.Font                 = New-Object System.Drawing.Font("Arial",13,[System.Drawing.FontStyle]::Regular)
$findButton.BackColor            = 'White'

$outputBox                       = New-Object system.Windows.Forms.RichTextBox
$outputBox.multiline             = $true
$outputBox.width                 = 400
$outputBox.height                = 370
$outputBox.location              = New-Object System.Drawing.Point(0,50)
$outputBox.BackColor             = 'WhiteSmoke'
$outputBox.Scrollbars            = "Vertical"
$outputBox.ReadOnly              = $true 
$outputBox.DetectUrls            = $false

$Button1                         = New-Object system.Windows.Forms.Button
$Button1.text                    = ""
$Button1.width                   = 120
$Button1.height                  = 30
$Button1.location                = New-Object System.Drawing.Point(10,430)
$Button1.BackColor               = 'WhiteSmoke'
$Button1.Enabled                 = $false
$Button1.Font                    = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Regular)

$Button2                         = New-Object system.Windows.Forms.Button
$Button2.text                    = ""
$Button2.width                   = 120
$Button2.height                  = 30
$Button2.location                = New-Object System.Drawing.Point(140,430)
$Button2.BackColor               = 'WhiteSmoke'
$Button2.Enabled                 = $false
$Button2.Font                    = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Regular)

$Button3                         = New-Object system.Windows.Forms.Button
$Button3.text                    = ""
$Button3.width                   = 120
$Button3.height                  = 30
$Button3.location                = New-Object System.Drawing.Point(270,430)
$Button3.BackColor               = 'WhiteSmoke'
$Button3.Enabled                 = $false
$Button3.Font                    = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Regular)

$Button4                         = New-Object system.Windows.Forms.Button
$Button4.text                    = ""
$Button4.width                   = 120
$Button4.height                  = 30
$Button4.location                = New-Object System.Drawing.Point(10,470)
$Button4.BackColor               = 'WhiteSmoke'
$Button4.Enabled                 = $false
$Button4.Font                    = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Regular)

$Button5                         = New-Object system.Windows.Forms.Button
$Button5.text                    = ""
$Button5.width                   = 120
$Button5.height                  = 30
$Button5.location                = New-Object System.Drawing.Point(140,470)
$Button5.BackColor               = 'WhiteSmoke'
$Button5.Enabled                 = $false
$Button5.Font                    = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Regular)

$Button6                         = New-Object system.Windows.Forms.Button
$Button6.text                    = ""
$Button6.width                   = 120
$Button6.height                  = 30
$Button6.location                = New-Object System.Drawing.Point(270,470)
$Button6.BackColor               = 'WhiteSmoke'
$Button6.Enabled                 = $false
$Button6.Font                    = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Regular)

$optionsButton                   = New-Object system.Windows.Forms.Button
$optionsButton.text              = ""
$optionsButton.width             = 125
$optionsButton.height            = 30
$optionsButton.location          = New-Object System.Drawing.Point(410,470)
$optionsButton.BackColor         = 'White'
$optionsButton.Text              = "Options"
$optionsButton.Font              = New-Object System.Drawing.Font("Arial",13,[System.Drawing.FontStyle]::Regular)

#this is called computerBox because it was intended for computers
#but as the script evolved it may now also be populated by AD-groups etc..
$computerBox                     = New-Object System.Windows.Forms.ComboBox
$computerBox.width               = 125
$computerBox.height              = 30
$computerBox.location            = New-Object System.Drawing.Point(410,430)
$computerBox.Font                = New-Object System.Drawing.Font("Arial",13,[System.Drawing.FontStyle]::Regular)
$computerBox.DropDownStyle       = [System.Windows.Forms.ComboBoxStyle]::DropDownList

$coverLabel1                     = New-Object system.Windows.Forms.Label
$coverLabel1.width               = 400
$coverLabel1.height              = 50
$coverLabel1.location            = New-Object System.Drawing.Point(0,0)
$coverLabel1.Font                = 'Microsoft Sans Serif,10'

$coverLabel2                     = New-Object system.Windows.Forms.Label
$coverLabel2.width               = 550
$coverLabel2.height              = 90
$coverLabel2.location            = New-Object System.Drawing.Point(0,420)
$coverLabel2.Font                = 'Microsoft Sans Serif,10'

$Form.controls.AddRange(@($inputBox,$objectBox,$findButton,$outputBox,$Button1,$Button2,$Button3,$Button4,$Button5,$Button6,$optionsButton,$computerBox,$coverLabel1,$coverLabel2))




