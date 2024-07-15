# Define network path and credentials
$networkPath = '\\it-server\Software-Install\Office365-New'
$username = 'berninausa\patchmgr.ca'
$password = 'gQ7uF@ysS!BY' | ConvertTo-SecureString -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($username, $password)

# Define setup command
$setupCommand = "setup.exe /configure Office365.xml"

# Define full path to setup.exe
$setupPath = Join-Path -Path $networkPath -ChildPath "setup.exe"

# Check if Office 365 is already installed
$officeInstalled = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -like "*Office 365*" }

if ($officeInstalled) {
    
    exit
}
else 
{
 # Run setup.exe with credentials, hiding all windows
    Start-Process -FilePath $setupPath -ArgumentList "/configure Office365.xml" -Credential $credential -wait -Windowstyle Hidden
}