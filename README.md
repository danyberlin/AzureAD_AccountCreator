# AzureAD AccountCreator
Simple PS Script to create AzureAD, MS365 user accounts with attributes in bulk by referencing an Excel workbook. The script can read/write the worksheet given as a parameter. 

Includes so far following attributes

- Licenses by SKU name
- Groups by Email or ObjectID, no **SharedMailbox**
- Company, Position, Department, Phone Number
- Display Name
- Requires only Email Adress for account creation or update
- Generated First and Lastname based on mandatoty Email pattern!
firstname.lastname@domainname (default domain name)


## Requirements ##
Windows Session by a domain user with the role User Administator or Global Administrator
PowerShell
`Install-Module -Name AzureAD`
`Set-ExecutionPolicy -ExecutionPolicy RemoteSigned`

## HowTo ##
1) Open the script
2) Make sure the static values match your worksheet's colum
Navigate to the folder where the script is located
Enter name of the script like a command with the following two parameters:
    name of the workbook (absolute path!) name of worksheet

eg: PS C:\Users\d\Desktop> azureConnect.ps1 C:\Users\d\Downloads\someWB.xlxs Sheet1

Good luck everyone!
D




