cls
# Static Values
$DisplayName_Pos = 2
$Email_Pos = 3
$License_Pos = 4
$Groups_Pos = 5
$Department_Pos = 6
$Position_Pos = 7
$Company_Pos = 8
$PhoneNumber_Pos = 9
$Password_Pos = 10
$Messages_Pos = 11

$path = $args[0] #spreadsheet
$CurrentWS = $args[1] #worksheet


# Helping Functions
function printNewcomer {
    param (
        $DisplayName, $EmailAddress, $Licences, $Groups, $Department, $Position, $FirstName, $LastName
    )
    Write-Output ""
    Write-Host -NoNewline "Newcomer: "
    Write-Host -ForegroundColor DarkGreen $DisplayName
    Write-Host -ForegroundColor DarkGreen "`t " $FirstName "`t " $LastName
    Write-Host -ForegroundColor DarkGreen "`t " $EmailAddress
    Write-Host -ForegroundColor DarkGreen "`t " $Licences
    Write-Host -ForegroundColor DarkGreen "`t " $Groups
    Write-Host -ForegroundColor DarkGreen "`t " $Position
    Write-Host -ForegroundColor DarkGreen "`t " $Department
    Write-Output ""
}
function Get-RandomCharacters($length, $characters) {
    $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length }
    $private:ofs = ""
    return [String]$characters[$random]
}
 
function Scramble-String([string]$inputString) {     
    $characterArray = $inputString.ToCharArray()   
    $scrambledStringArray = $characterArray | Get-Random -Count $characterArray.Length     
    $outputString = -join $scrambledStringArray
    return $outputString 
}
function generatePW {
    $password = Get-RandomCharacters -length 1 -characters 'ABCDEFGHKLMNOPRSTUVWXYZ'
    $password += Get-RandomCharacters -length 3 -characters 'abcdefghiklmnoprstuvwxyz'
    $password += Get-RandomCharacters -length 3 -characters '1234567890'
    $password += Get-RandomCharacters -length 1 -characters '!$%&=?@#+'
    # $password = Scramble-String $password
    return $password
}


Write-Output "Creating Excel.Application Object"
$excel = new-object -comobject Excel.Application # create base object
$excel.visible = $true # make Excel visible
$excel.DisplayAlerts = $false;
$excel.WindowState = "xlMaximized"

Write-Output "Opening Workbook"
try {
    $wb = $excel.workbooks.open($path)    # open the workbook
}
catch {
    Write-Output $_.Exception.Message
    Write-Output "Closing Excel"
    $excel.Quit()
    throw $_
}
try {
    Write-Output "Opening Worksheet"
    $ws = $wb.Worksheets.item($CurrentWS)    # open worksheet

}
catch {
    Write-Output $_.Exception.Message
    Write-Output "Closing Workbook"
    $wb.Close()
    Write-Output "Closing Excel"
    $excel.Quit()
    throw $_
}

$newcomerCount = ($ws.UsedRange.Rows).count - 1
Write-Output ""
Write-Host -NoNewline "Total Number of Newcomer: "
Write-Host -ForegroundColor DarkGreen $newcomerCount
Write-Output ""
Write-Output "Does that make sense? Ctr+C to quit or Any Key to continue..."
Write-Output "CHECK FOR ENOUGH LICENES!!!"
$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
Write-Output ""
# Workbook open, work begins! 

# Alternative authentication when user account is not administrator but uses different admin credentials

# if ($null -eq $AzureAdCred) {
#     $AzureAdCred = Get-Credential
# }
# Connect-AzureAD -Credential $AzureAdCred

#Authentication
$UPN = whoami /upn
Connect-AzureAD -AccountId $UPN


$keepGoing = $true
while ($keepGoing) {
    

    for ($i = 1; $i -lt $newcomerCount + 1; $i++) {
        if (($ws.Cells.Item($i + 1, $Messages_Pos)).Text -eq "OK" -or ($ws.Cells.Item($i + 1, $Messages_Pos)).Text -eq "SKIP") {
            continue
        }
        Write-Host -BackgroundColor Yellow "-------------------------------------"
        Write-Output ""
        $ws.Cells.Item($i + 1, $Messages_Pos) = "In Process"
        $ws.Cells.Item($i + 1, $Messages_Pos).Interior.ColorIndex = 44
        $DisplayName = $ws.Cells.Item($i + 1, $DisplayName_Pos).Text
        $EmailAddress = $ws.Cells.Item($i + 1, $Email_Pos).Text
        
        # Illegal characters
        # "ä","ü","ö","ß"," ","å","é","®","þ","í","ó","ß","ð","â"

        if($ws.Cells.Item($i + 1, $Email_Pos).Text.contains("ä") -or $ws.Cells.Item($i + 1, $Email_Pos).Text.contains("ü") -or $ws.Cells.Item($i + 1, $Email_Pos).Text.contains("ö") -or $ws.Cells.Item($i + 1, $Email_Pos).Text.contains("ß") -or $ws.Cells.Item($i + 1, $Email_Pos).Text.contains(" ") -or $ws.Cells.Item($i + 1, $Email_Pos).Text.contains("å") -or $ws.Cells.Item($i + 1, $Email_Pos).Text.contains("é") -or $ws.Cells.Item($i + 1, $Email_Pos).Text.contains("®") -or $ws.Cells.Item($i + 1, $Email_Pos).Text.contains("þ") -or $ws.Cells.Item($i + 1, $Email_Pos).Text.contains("í") -or $ws.Cells.Item($i + 1, $Email_Pos).Text.contains("ó") -or $ws.Cells.Item($i + 1, $Email_Pos).Text.contains("ð") -or $ws.Cells.Item($i + 1, $Email_Pos).Text.contains("â")) {
           $ws.Cells.Item($i + 1, $Email_Pos).Interior.ColorIndex = 3
           $ws.Cells.Item($i + 1, $Messages_Pos) = "EMAIL CONTAINS ILLEGAL CHARACTERS"
           continue
        }

        $EmailNickName = ($ws.Cells.Item($i + 1, $Email_Pos).Text).split("@")
        $Name = ($EmailNickName).split(".")
        $TextInfo = (Get-Culture).TextInfo
        $FirstName = $TextInfo.ToTitleCase($Name[0])
        $LastName = $TextInfo.ToTitleCase($Name[1])
        $Licences = ($ws.Cells.Item($i + 1, $License_Pos).Text).split(",")
        $Groups = ($ws.Cells.Item($i + 1, $Groups_Pos).Text).split(",")
        $Department = $ws.Cells.Item($i + 1, $Department_Pos).Text
        $Position = $ws.Cells.Item($i + 1, $Position_Pos).Text
        $Company = $ws.Cells.Item($i + 1, $Company_Pos).Text
        $PhoneNumber = $ws.Cells.Item($i + 1, $PhoneNumber_Pos).Text
        if (!$EmailAddress) {
            $ws.Cells.Item($i + 1, $Messages_Pos) = "Missing Value!"
            $ws.Cells.Item($i + 1, $Messages_Pos).Interior.ColorIndex = 9
            continue
        } 
        printNewcomer $DisplayName $EmailAddress $Licences $Groups $Department $Position $FirstName $LastName
        $pw = generatePW
        # Start creating accounts!

        # Define a Password Profile
        $PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
        $PasswordProfile.EnforceChangePasswordPolicy = 0
        $PasswordProfile.ForceChangePasswordNextLogin = 0
        $PasswordProfile.Password = $pw

        Try {
            # Check if user exists
            Write-Host -NoNewline "Checking if user "
            Write-Host -NoNewline -ForegroundColor DarkGreen $EmailAddress
            Write-Host " exists"
            Write-Output ""

            try {
                $userCheck = Get-AzureADUser -ObjectId $EmailAddress
            }
            catch {
                Write-Output ""
            }

            if ($null -eq $userCheck) {
                Write-Output "User not found. Continue"

                # Creating new user with basic attributes
                New-AzureADUser -DisplayName $DisplayName -AccountEnabled $true -UserPrincipalName $EmailAddress -MailNickName $EmailNickName[0] -GivenName $FirstName -Surname $LastName -PasswordProfile $PasswordProfile -UsageLocation "DE" | Out-Null
                Write-Verbose "$DisplayName : AAD Account created successfully!"
                $ws.Cells.Item($i + 1, $Email_Pos).Interior.ColorIndex = 43
                $ws.Cells.Item($i + 1, $Password_Pos) = $pw
      
            }
            else {
                Write-Host -ForegroundColor Yellow "User already exists!"
                Write-Output "Trying to update or add information..."
                Write-Output ""
            }
       
            # Adding additional info (Department, Position)
            $ADuser = Get-AzureADUser -ObjectID "$EmailAddress"
            if ($Position) {
                try {
                    Write-Output "Adding Position: $Position"
                    Set-AzureADUser -ObjectId $ADuser.ObjectID -JobTitle $Position
                    $ws.Cells.Item($i + 1, $Position_Pos).Interior.ColorIndex = 43
                }
                catch {
                    Write-Error "Error Adding Position: $_"
                }
            }
            else {
                Write-Output "No Position to Add"
            }

            if ($Department) {
                try {
                    Write-Output "Adding Department $Department"
                    Set-AzureADUser -ObjectId $ADuser.ObjectID -Department $Department
                    $ws.Cells.Item($i + 1, $Department_Pos).Interior.ColorIndex = 43
                }
                catch {
                    Write-Error "Error Adding Department: $_"
                }
            }
            else {
                Write-Output "No Department to Add"
            }

            if ($Company) {
                try {
                    Write-Output "Adding Company $Company"
                    Set-AzureADUser -ObjectId $ADuser.ObjectID -CompanyName $Company
                    $ws.Cells.Item($i + 1, $Company_Pos).Interior.ColorIndex = 43
                }
                catch {
                    Write-Error "Error Adding Company: $_"
                }
            }
            else {
                Write-Output "No Company to Add"
            }

            # Might use this one day
            if ($PhoneNumber) {
                try {
                    Write-Output "Adding Phonenumber $PhoneNumber"
                    Set-AzureADUser -ObjectId $ADuser.ObjectID -TelephoneNumber $PhoneNumber
                    $ws.Cells.Item($i + 1, $PhoneNumber).Interior.ColorIndex = 43
                }
                catch {
                    Write-Host "Some Error Adding Phone Number but probably not really" 
                }
            }
            else {
                Write-Output "No Phone Number to Add"
            }

            # Adding all groups
            if ($Groups) {
                for ($j = 0; $j -lt $Groups.Count; $j++) {
                    $GroupName = $Groups[$j]
                    $ActualGroupName = "$GroupName"
                    try {
                        Write-Host -NoNewline "Searching the Group: $ActualGroupName"
                        $AADGroupID = Get-AzureADGroup -SearchString "$ActualGroupName"
                        Write-Host -ForegroundColor DarkGreen " ...OK"
                        if ($AADGroupID.Length -eq 0) {
                            $AADGroupID = Get-AzureADGroup -ObjectId "$ActualGroupName"
                        }
                        Write-Output ""
                    
                    }
                    catch {
                        Write-Error "$AADGroupID : does not exist. $_"
                    } 
                    try {
                        Write-Host -NoNewline "Assigning $EmailAddress to " $ActualGroupName
                        $ADuser = Get-AzureADUser -ObjectID "$EmailAddress"
                        Add-AzureADGroupMember -ObjectID $AADGroupID.ObjectID -RefObjectId $ADuser.ObjectID
                        $ws.Cells.Item($i + 1, $Groups_Pos).Interior.ColorIndex = 43
                        Write-Host -ForegroundColor DarkGreen " ...OK"
                        Write-Output ""
                    }
                    catch {
                        Write-Output ""
                        Write-Host -ForegroundColor Red "Error In Groups: " $Groups[$j]
                        Write-Output "Possibly Already a member, check carefully"
                    }
                    Write-Output ""
                }  
            }
            else {
                Write-Output "No Groups to Add"
            }
      
        
            # Adding all Licenses
            if ($Licences) {
                for ($k = 0; $k -lt $Licences.Count; $k++) {
                    try {
                        Write-Host -NoNewline "Assigning License: " $Licences[$k]
                        $License = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
                        $License.SkuId = (Get-AzureADSubscribedSku | Where-Object -Property SkuPartNumber -Value $Licences[$k] -EQ).SkuID
                        $LicensesToAssign = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
                        $LicensesToAssign.AddLicenses = $License
                        Set-AzureADUserLicense -ObjectId $EmailAddress -AssignedLicenses $LicensesToAssign
                        $ws.Cells.Item($i + 1, $License_Pos).Interior.ColorIndex = 43
                        Write-Host -ForegroundColor DarkGreen " ...OK"
                        Write-Output ""
                    }
                    catch {
                        Write-Host "Error Adding License: " $Licences[$k]
                        Write-Error $_
                    }
                }  
            }
            else {
                Write-Output "No Licenses to Add"
            }
            
            # Setting log column as OK to say everything went well
            $ws.Cells.Item($i + 1, $Messages_Pos) = "OK"
            $ws.Cells.Item($i + 1, $Messages_Pos).Interior.ColorIndex = 43
        }
        Catch {
            Write-Error "$DisplayName : An Error occurred while creating AAD Account. $_"
            $ws.Cells.Item($i + 1, $Messages_Pos) = "ERROR"
            $ws.Cells.Item($i + 1, $Messages_Pos).Interior.ColorIndex = 46
        }
        Write-Host -ForegroundColor Green "All Good"
    }
    $finalQuestion = Read-Host "Enter c to go again or any key to quit"
    
    $newcomerCount = ($ws.UsedRange.Rows).count - 1
    Write-Output ""
    Write-Host -NoNewline "Total Number of Newcomer: "
    Write-Host -ForegroundColor DarkGreen $newcomerCount
    Write-Output ""
    Write-Output "Does that make sense? Ctr+C to quit or Any Key to continue..."
    Write-Output "CHECK FOR ENOUGH LICENES!!!"
    $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    if ($finalQuestion -ne "c") {
        $keepGoing = $false
    }
}


$excel.ActiveWorkbook.Save();

Write-Output ""
Write-Host -BackgroundColor Yellow "-------------------------------------"
Write-Output ""
Write-Output "Closing Workbook"
$wb.Close()
Write-Output "Closing Excel"
$excel.Quit()

