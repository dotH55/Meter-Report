### Elevate User
param([switch]$Elevated)
Add-Type -AssemblyName System.IO.Compression.FileSystem

### Function for unzipping a file
function Unzip
{
    param([string]$zipfile, [string]$outpath)

    [System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile, $outpath)
}

### Credentials
$USERNAME = ''
$PASSWORD = ''

### Variables
$DATE = get-date -format 'yyyy-MM-dd'
$FILE_NAME = 'US_BILLING_' + $DATE + '_01.zip'
$URL = 'https://' + $DATE + ''

$DATE = get-date -format 'yyyyMMdd'
$ZIP_FILE = 'US_BILLING_' + $DATE + '_01.zip'

### Browser
$IE = new-object -com "InternetExplorer.Application"
### Visibility  == False
$IE.visible = $false
$IE.navigate($URL)

### Authentification
#Start-Sleep -seconds 1
$SHELL = New-Object -ComObject wscript.shell;
If ($SHELL.AppActivate('Windows Security') -eq $false){Start-Sleep -seconds 1}
$SHELL.SendKeys($USERNAME)
start-sleep -seconds 1
$SHELL.SendKeys('{TAB}')
start-sleep -seconds 1
$SHELL.SendKeys($PASSWORD)
start-sleep -seconds 1
$SHELL.SendKeys('{TAB}{TAB}')
$SHELL.SendKeys('{ENTER}')

### Get Zip_File Url
while ($IE.Busy -eq $true){Start-Sleep -seconds 1;}
$INNER_TEXT = $IE.Document.IHTMLDocument2_body.innerText
$X = $INNER_TEXT.IndexOf('https://')
$URL = $INNER_TEXT.Substring($X)

### Download
wget $URL -OutFile $ZIP_FILE
start-sleep -seconds 3

### Unzip
Unzip $ZIP_FILE .\
Start-Sleep -seconds 2

python .\PY_Script.pyw

### Remove
rm $ZIP_FILE

### Terminate
Exit(0)