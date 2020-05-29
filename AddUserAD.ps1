#Matt Smith 2020

Add-Type -AssemblyName System.Windows.Forms
Add-type -AssemblyName System.Web
[System.Windows.Forms.Application]::EnableVisualStyles()

$AddUserAD = New-Object system.Windows.Forms.Form
$AddUserAD.ClientSize = '400,500'
$AddUserAD.text = "AddUserAD"
$AddUserAD.TopMost = $false

$FirstName1 = New-Object system.Windows.Forms.TextBox
$FirstName1.multiline = $false
$FirstName1.width = 100
$FirstName1.height = 20
$FirstName1.location = New-Object System.Drawing.Point(92, 18)
$FirstName1.Font = 'Microsoft Sans Serif,10'

$LastName1 = New-Object system.Windows.Forms.TextBox
$LastName1.multiline = $false
$LastName1.width = 122
$LastName1.height = 20
$LastName1.location = New-Object System.Drawing.Point(265, 18)
$LastName1.Font = 'Microsoft Sans Serif,10'

$FirstName2 = New-Object system.Windows.Forms.TextBox
$FirstName2.multiline = $false
$FirstName2.width = 100
$FirstName2.height = 20
$FirstName2.location = New-Object System.Drawing.Point(92, 46)
$FirstName2.Font = 'Microsoft Sans Serif,10'

$LastName2 = New-Object system.Windows.Forms.TextBox
$LastName2.multiline = $false
$LastName2.width = 122
$LastName2.height = 20
$LastName2.location = New-Object System.Drawing.Point(265, 46)
$LastName2.Font = 'Microsoft Sans Serif,10'

$Username = New-Object system.Windows.Forms.TextBox
$Username.multiline = $false
$Username.width = 135
$Username.height = 20
$Username.location = New-Object System.Drawing.Point(16, 77)
$Username.Font = 'Microsoft Sans Serif,10'
$Username.Enabled = $false
$Username.TextAlign = 'Right'
$AddUserAD.Controls.Add($Username) #activating the text box inside the primary window
$Username.Text = ""

$Domain = New-Object system.Windows.Forms.ComboBox
$Domain.text = " "
$Domain.width = 148
$Domain.height = 20
$Domain.location = New-Object System.Drawing.Point(175, 77)
$Domain.Font = 'Microsoft Sans Serif,10'
$DomainAD = get-aduser -Filter * -property UserPrincipalName | where { $_.UserPrincipalName -ne $null } | % { $_.UserPrincipalName.Split('@')[1] } | sort-object  -unique
$Domain.Items.AddRange($DomainAD)
$Domain.SelectedIndex = 1

$Department = New-Object system.Windows.Forms.ComboBox
$Department.text = " "
$Department.width = 294
$Department.height = 20
$Department.location = New-Object System.Drawing.Point(92, 121)
$Department.Font = 'Microsoft Sans Serif,10'
$DepartmentsAD = get-aduser -filter * -property department | select -ExpandProperty department | sort-object  -unique
$Department.Items.AddRange($DepartmentsAD)

$Title = New-Object system.Windows.Forms.ComboBox
$Title.text = " "
$Title.width = 295
$Title.height = 20
$Title.location = New-Object System.Drawing.Point(92, 150)
$Title.Font = 'Microsoft Sans Serif,10'

$Manager = New-Object system.Windows.Forms.ComboBox
$Manager.text = " "
$Manager.width = 295
$Manager.height = 20
$Manager.location = New-Object System.Drawing.Point(92, 180)
$Manager.Font = 'Microsoft Sans Serif,10'

$Extention = New-Object system.Windows.Forms.TextBox
$Extention.multiline = $false
$Extention.width = 100
$Extention.height = 20
$Extention.location = New-Object System.Drawing.Point(92, 210)
$Extention.Font = 'Microsoft Sans Serif,10' 

$Password = New-Object system.Windows.Forms.TextBox
$Password.multiline = $false
$Password.width = 210
$Password.height = 20
$Password.location = New-Object System.Drawing.Point(92, 240)
$Password.Font = 'Microsoft Sans Serif,10'

$UserProp = New-Object system.Windows.Forms.TextBox
$UserProp.multiline = $true
$UserProp.ScrollBars = "Vertical"
$UserProp.width = 370
$UserProp.height = 150
$UserProp.location = New-Object System.Drawing.Point(16, 290)
$UserProp.Font = 'Microsoft Sans Serif,10'
$UserProp.Enabled = $false
$AddUserAD.Controls.Add($UserProp) #activating the text box inside the primary window
$UserProp.Text = " "

$FirstName1Label = New-Object system.Windows.Forms.Label
$FirstName1Label.text = "First Name"
$FirstName1Label.AutoSize = $true
$FirstName1Label.width = 25
$FirstName1Label.height = 10
$FirstName1Label.location = New-Object System.Drawing.Point(16, 22)
$FirstName1Label.Font = 'Microsoft Sans Serif,10'

$FirstName2Label = New-Object system.Windows.Forms.Label
$FirstName2Label.text = "Confirm"
$FirstName2Label.AutoSize = $true
$FirstName2Label.width = 25
$FirstName2Label.height = 10
$FirstName2Label.location = New-Object System.Drawing.Point(16, 50)
$FirstName2Label.Font = 'Microsoft Sans Serif,10'

$LastName1Label = New-Object system.Windows.Forms.Label
$LastName1Label.text = "Last Name"
$LastName1Label.AutoSize = $true
$LastName1Label.width = 25
$LastName1Label.height = 10
$LastName1Label.location = New-Object System.Drawing.Point(195, 22)
$LastName1Label.Font = 'Microsoft Sans Serif,10'

$LastName2Label = New-Object system.Windows.Forms.Label
$LastName2Label.text = "Confirm"
$LastName2Label.AutoSize = $true
$LastName2Label.width = 25
$LastName2Label.height = 10
$LastName2Label.location = New-Object System.Drawing.Point(195, 50)
$LastName2Label.Font = 'Microsoft Sans Serif,10'

$DepartmentLabel = New-Object system.Windows.Forms.Label
$DepartmentLabel.text = "Department"
$DepartmentLabel.AutoSize = $true
$DepartmentLabel.width = 25
$DepartmentLabel.height = 10
$DepartmentLabel.location = New-Object System.Drawing.Point(16, 125)
$DepartmentLabel.Font = 'Microsoft Sans Serif,10'

$TitleLabel = New-Object system.Windows.Forms.Label
$TitleLabel.text = "Title"
$TitleLabel.AutoSize = $true
$TitleLabel.width = 25
$TitleLabel.height = 10
$TitleLabel.location = New-Object System.Drawing.Point(16, 152)
$TitleLabel.Font = 'Microsoft Sans Serif,10'

$ManagerLabel = New-Object system.Windows.Forms.Label
$ManagerLabel.text = "Manager"
$ManagerLabel.AutoSize = $true
$ManagerLabel.width = 25
$ManagerLabel.height = 10
$ManagerLabel.location = New-Object System.Drawing.Point(16, 184)
$ManagerLabel.Font = 'Microsoft Sans Serif,10'

$ExtentionLabel = New-Object system.Windows.Forms.Label
$ExtentionLabel.text = "Extention"
$ExtentionLabel.AutoSize = $true
$ExtentionLabel.width = 25
$ExtentionLabel.height = 10
$ExtentionLabel.location = New-Object System.Drawing.Point(16, 215)
$ExtentionLabel.Font = 'Microsoft Sans Serif,10'

$PasswordLabel = New-Object system.Windows.Forms.Label
$PasswordLabel.text = "Password"
$PasswordLabel.AutoSize = $true
$PasswordLabel.width = 25
$PasswordLabel.height = 10
$PasswordLabel.location = New-Object System.Drawing.Point(16, 245)
$PasswordLabel.Font = 'Microsoft Sans Serif,10'

$ADSync = New-Object system.Windows.Forms.Label
$ADSync.text = "AD Last Sync: -- Min"
$ADSync.AutoSize = $true
$ADSync.width = 25
$ADSync.height = 10
$ADSync.location = New-Object System.Drawing.Point(235, 445)
$ADSync.Font = 'Microsoft Sans Serif,10'

$VerifyBox = New-Object system.Windows.Forms.CheckBox
$VerifyBox.text = "All of this information is correct"
$VerifyBox.AutoSize = $false
$VerifyBox.width = 220
$VerifyBox.height = 20
$VerifyBox.Anchor = 'bottom,left'
$VerifyBox.location = New-Object System.Drawing.Point(16, 443)
$VerifyBox.Font = 'Microsoft Sans Serif,10'

$ProgressBar1 = New-Object system.Windows.Forms.ProgressBar
$ProgressBar1.width = 181
$ProgressBar1.height = 29
$ProgressBar1.Anchor = 'bottom,left'
$ProgressBar1.location = New-Object System.Drawing.Point(15, 464)
$ProgressBar1.text = "Check Fields"
$Progressbar1.Maximum = 4
$Progressbar1.Step = 1
$Progressbar1.Value = 0

$AddUserButton = New-Object system.Windows.Forms.Button
$AddUserButton.text = "Create User"
$AddUserButton.width = 185
$AddUserButton.height = 29
$AddUserButton.Anchor = 'right,bottom'
$AddUserButton.location = New-Object System.Drawing.Point(204, 464)
$AddUserButton.Font = 'Microsoft Sans Serif,10'

$ExtentionButton = New-Object system.Windows.Forms.Button
$ExtentionButton.text = "Generate"
$ExtentionButton.width = 80
$ExtentionButton.height = 25
$ExtentionButton.location = New-Object System.Drawing.Point(308, 210)
$ExtentionButton.Font = 'Microsoft Sans Serif,10'

$PasswordButton = New-Object system.Windows.Forms.Button
$PasswordButton.text = "Generate"
$PasswordButton.width = 80
$PasswordButton.height = 25
$PasswordButton.location = New-Object System.Drawing.Point(308, 239)
$PasswordButton.Font = 'Microsoft Sans Serif,10'

$UsernameFillButton = New-Object system.Windows.Forms.Button
$UsernameFillButton.text = "Check"
$UsernameFillButton.width = 60
$UsernameFillButton.height = 26
$UsernameFillButton.location = New-Object System.Drawing.Point(326, 76)
$UsernameFillButton.Font = 'Microsoft Sans Serif,10'

$UserPropButton = New-Object system.Windows.Forms.Button
$UserPropButton.text = "⟳"
$UserPropButton.width = 20
$UserPropButton.height = 20
$UserPropButton.location = New-Object System.Drawing.Point(369, 443)
$UserPropButton.Font = 'Microsoft Sans Serif,10'

$AtSignLabel = New-Object system.Windows.Forms.Label
$AtSignLabel.text = "@"
$AtSignLabel.AutoSize = $true
$AtSignLabel.width = 26
$AtSignLabel.height = 10
$AtSignLabel.location = New-Object System.Drawing.Point(155, 82)
$AtSignLabel.Font = 'Microsoft Sans Serif,10'

$AddUserAD.controls.AddRange(@($FirstName1, $LastName1, $FirstName2, $LastName2, $Username, $Domain, $Department, $Title, $FirstName1Label, $FirstName2Label, $LastName1Label, $LastName2Label, $DepartmentLabel, $TitleLabel, $VerifyBox, $ProgressBar1, $AddUserButton, $UsernameFillButton, $AtSignLabel, $UserProp, $Manager, $ManagerLabel, $ExtentionLabel, $Extention, $UserPropButton, $PasswordLabel, $Password, $PasswordButton, $ExtentionButton, $ADSync))


$UsernameFillButton.add_Click( { CheckUsername })
$AddUserButton.add_Click( { createUser })
$UserPropButton.add_Click( { WelcomeLetter })
#$UserPropButton.add_Click( { UserProp })
$PasswordButton.add_Click( { GenPass })
$ExtentionButton.add_Click( { GenEx })



$Department_SelectedIndexChanged = {
    $Title.Items.Clear()
    $TitleAD = get-aduser -Filter { Department -eq $Department.Text } -properties department, title | select -ExpandProperty title | sort-object  -unique
    $Title.Items.AddRange($TitleAD)
}

$Title_SelectedIndexChanged = {
    $Manager.Items.Clear()
    $ManagerAD = get-aduser -Filter { Title -eq $Title.Text } -properties Manager | % { $_.Manager.Split('='',')[1] } | sort-object  -unique
    $Manager.Items.AddRange($ManagerAD)
    $Manager.SelectedIndex = 0
}



$FirstName1_TextChanged = {

    $Username.Text = " "
    IF (Compare-Object $FirstName1 $FirstName2) {
        $Global:UsernameGlobalVar = "Check First Name"
        $Username.AppendText($Global:UsernameGlobalVar)
        return
    }
    IF (Compare-Object $LastName1 $LastName2) {
        $Global:UsernameGlobalVar = "Check Last Name"
        $Username.AppendText($Global:UsernameGlobalVar)
        return
    }
    ELSE {
        Set-Variable -name UsernameGlobalVar -Scope Global -Value 'POPULATED'
        $Global:UsernameGlobalVar = "$($FirstName1.Text.SubString(0,1))$($LastName1.Text)"
        $Username.Clear()
        $Username.AppendText($Global:UsernameGlobalVar)
        return
    }
}


function CheckUsername() {
    $UsernameEmail = "$($UsernameGlobalVar)@$($Domain.Text)"
    if (get-aduser -filter "samaccountname -eq '$($Username.Text)'") {
        $Username.ForeColor = 'Red'
        $Username.Enabled = $true
    }
    else {
        $Username.ForeColor = 'Green'
        $Username.Enabled = $true
    }
}

function GenEx() {
    $Extention.Text = "$(Get-Random -Minimum 3000 -Maximum 3999)"
    if ($ExtentionAD.Contains($Extention.Text)) {
        GenEx
    } 
}

function GenPass() {
    $Password.Text = "Trop$("@", "!", "#", "$", "%", "?" | Get-Random)$(Get-Random -Minimum 0000 -Maximum 9999)"
}

function UserProp() {
    Connect-MsolService

    $LastSync = Get-MsolCompanyInformation | % { $_.LastDirSyncTime.AddHours(-5) }
    $CurrentTime = (GET-DATE)
    $LastSyncMin = NEW-TIMESPAN –Start $LastSync –End $CurrentTime | % { $_.TotalMinutes }
    $LastSyncMin = [math]::Round($LastSyncMin, 0)
    $ADSync.Text = "AD Last Sync: $($LastSyncMin) Min"
}

$ExtentionAD = get-aduser -filter * -property telephoneNumber | where { $_.telephoneNumber -ne $null } | % { $_.telephoneNumber.Split('-''-')[2] } | sort-object  -unique
$Extention_TextChanged = {
    $Extention.ForeColor = 'Red'
    if ( $Extention.Text -match "^\d+$") {
        $Extention.ForeColor = 'Green'
    }
    if ( $ExtentionAD.Contains($Extention.Text)) {
        $Extention.ForeColor = 'Red'
    }
    if ( 1000 -gt $Extention.Text) {
        $Extention.ForeColor = 'Red'
    }
    if ( 9999 -lt $Extention.Text) {
        $Extention.ForeColor = 'Red'
    }
}

#Connect-MsolService
function ADSyncUpdater($ADSync) {
    $LastSync = Get-MsolCompanyInformation | % { $_.LastDirSyncTime.AddHours(-5) }
    $CurrentTime = (GET-DATE)
    $LastSyncMin = NEW-TIMESPAN –Start $LastSync –End $CurrentTime | % { $_.TotalMinutes }
    $LastSyncMin = [math]::Round($LastSyncMin, 0)
    $ADSync.Text = "AD Last Sync: $($LastSyncMin) Min"
}

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 10000
$timer.Add_Tick( { ADSyncUpdater $ADSync })
$timer.Enabled = $True
$timer.Dispose()


function createUser() {

    $UserProp.Clear()

    If ($VerifyBox.Checked -eq $true) {
        $ProgressBar1.PerformStep();

        #Get more info and set variables
        $titlevar = $Title.Text
        $managervar = $Manager.Text
        $departmentvar = $Department.Text
        $ManagerUsername = get-aduser -Filter { displayname -eq $managervar -and department -eq $departmentvar } -properties * | % { $_.SamAccountName } | sort-object  -unique
        $StreetAddress = get-aduser -Filter { Title -eq $titlevar } -properties StreetAddress | % { $_.StreetAddress } | sort-object  -unique
        $City = get-aduser -Filter { Title -eq $titlevar } -properties City | % { $_.City } | sort-object  -unique
        $State = get-aduser -Filter { Title -eq $titlevar } -properties State | % { $_.State } | sort-object  -unique
        $PostalCode = get-aduser -Filter { Title -eq $titlevar } -properties PostalCode | % { $_.PostalCode } | sort-object  -unique
        $Country = get-aduser -Filter { Title -eq $titlevar } -properties Country | % { $_.Country } | sort-object  -unique
        $Company = get-aduser -Filter { Title -eq $titlevar } -properties Company | % { $_.Company } | sort-object  -unique
        $fullpath = get-aduser -Filter { Title -eq $titlevar } -properties DistinguishedName | % { $_.DistinguishedName.Substring($_.DistinguishedName.IndexOf(',') + 1) } | sort-object  -unique
        $usernamevar = $Username.Text
        $firstnamevar = $Firstname1.Text
        $lastnamevar = $Lastname1.Text
        $departmentvar = $Department.Text
        $extentionvar = $Extention.Text
        $displaynamevar = "$firstnamevar $lastnamevar"
        $homedirectoryvar = "\\Fileserver01\User$\$usernamevar"
        $passvar = $Password.Text
        $SecurePass = (ConvertTo-SecureString $Password.Text -AsPlainText -Force)

        if ($null -eq $ManagerUsername){
            $UserProp.AppendText("Manager not found, please check department 'n")
        }

        $ProgressBar1.PerformStep();

        #Create account 
        $UsernameEmail = "$($Username.Text)@$($Domain.Text)"
        rm newuser.txt
        New-Item -Path . -Name "newuser.txt" -ItemType "file" -Value "New-ADUser 
-SamAccountName '$($Username.Text)' 
-GivenName '$($FirstName1.Text)' 
-Surname '$($LastName1.Text)' 
-Name '$($FirstName1.Text) $($LastName1.Text)' 
-DisplayName '$($FirstName1.Text) $($LastName1.Text)' 
-Description '$($Title.Text)' 
-EmailAddress '$UsernameEmail' 
-StreetAddress  '$StreetAddress' 
-City  '$City' 
-State  '$State' 
-PostalCode  '$PostalCode' 
-Country  '$Country' 
-ScriptPath Slogic.bat 
-HomeDrive 'U:' 
-HomeDirectory '\\Fileserver01\User$\$($Username.Text)' 
-OfficePhone $($Extention.Text) 
-Title '$($Title.Text)' 
-Department '$($Department.Text)' 
-Company '$($Company.Text)' 
-Manager '$ManagerUsername' 
-AccountPassword $(ConvertTo-SecureString $Password.Text -AsPlainText -Force) 
-Path $fullpath
-HomePage 'www.google.com'"

        New-ADUser -SamAccountName $usernamevar -GivenName $firstnamevar -Surname $lastnamevar -Name $displaynamevar -DisplayName $displaynamevar -Description $titlevar -EmailAddress $UsernameEmail -StreetAddress  $StreetAddress -City  $City -State  $State -PostalCode  $PostalCode -Country  $Country -ScriptPath 'Slogic.bat' -HomeDrive 'U:' -HomeDirectory $homedirectoryvar -Title $titlevar -Department $departmentvar -Company $Company -Manager $ManagerUsername -AccountPassword $SecurePass -Path $fullpath -UserPrincipalName $UsernameEmail -HomePage "www.google.com" -Enabled $True -OtherAttributes @{'ipPhone' = $extentionvar }

        $ProgressBar1.PerformStep();

        #Add MemberOf to new account
        $Groups = (get-aduser -Filter { Title -eq $titlevar } -properties * | Select-Object MemberOf).MemberOf | Get-ADGroup -Server :3268 | sort-object  -unique
        foreach ($Group in $Groups) {
            Write-Output $Group.Name
            Add-Content newuser.txt "Add-ADGroupMember -Identity '$($Group.Name)' -Members '$($Username.Text)'"
            $groupnamevar = $Group.Name
            Add-ADGroupMember -Identity $groupnamevar -Members $usernamevar
        }

        $ProgressBar1.PerformStep();
        $UserProp.AppendText("The user '$displaynamevar' has successfully been created")

    }
    else {
        $UserProp.AppendText("Verify information is correct")
    }



}

function WelcomeLetter1{
    $word = New-Object -comobject word.application
    $word.visible = $true
    $doc = $word.documents.open("WelcomeLetterTemplate.docx")

    $bookmark = $doc.Bookmarks("username")
    $bookmark.Range.Text = $usernamevar
    $bookmark = $doc.Bookmarks("password")
    $bookmark.Range.Text = $passvar
    $bookmark = $doc.Bookmarks("email")
    $bookmark.Range.Text = $UsernameEmail

}

Function WelcomeLetter($passvar,$UsernameEmail,$usernamevar)
{
    $word = New-Object -ComObject "Word.application"
    $doc = $word.Documents.Open("\\fs01\Softwares\NewUserAD\bin\WelcomeLetterTemplate.docx")
    $selection = $word.Selection
    
    $doc.Bookmarks.Add("username",$filluser).Range
    $filluser.Text = $UsernameGlobalVar
    
    $doc.Bookmarks.Add("password",$fillpass).Range
    $fillpass.Text = $passvar
    
    $doc.Bookmarks.Add("Email",$fillemail).Range
    $fillemail.Text = $usernamevar
    
    
            
    $file = "test1.docx"
    $doc.SaveAs([ref]$file)
    
    $Word.Quit()
}


$Department.add_SelectedIndexChanged($Department_SelectedIndexChanged)
$Title.add_SelectedIndexChanged($Title_SelectedIndexChanged)
$FirstName1.add_TextChanged($FirstName1_TextChanged)
$FirstName2.add_TextChanged($FirstName1_TextChanged)
$LastName1.add_TextChanged($FirstName1_TextChanged)
$LastName2.add_TextChanged($FirstName1_TextChanged)
$Extention.add_TextChanged($Extention_TextChanged)


[void]$AddUserAD.ShowDialog()
