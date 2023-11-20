function Test-RegistrySetting {
    param (
        [string]$app,
        [string]$path,
        [string]$key,
        [int]$expectedValue,
        [string]$message,
        [bool]$isCommonPath = $true
    )

    if ($isCommonPath -eq $true){
    $fullPath = "Registry::HKCU\Software\Policies\Microsoft\Cloud\Office\16.0\$app\$path"
    }
    else { $fullPath = "Registry::HKCU\Software\Policies\Microsoft\Cloud\Office\$($path)" }
    
    $regValue = (Get-ItemProperty -Path $fullPath -Name $key -ErrorAction SilentlyContinue).$key

    # Debugging output
    Write-Debug "Checking $app ($fullPath) for ${key}: Value = $regValue, Expected = $expectedValue"

    if ($regValue -ne $expectedValue) {
        Write-Output "$($app) - The Setting of - '$($message)' is incorrect"
    }
}

$settings = @(
    @{ Apps = @("Access", "Excel", "Powerpoint", "Publisher", "Visio", "Word"); Path = "security"; Key = "vbawarnings"; ExpectedValue = 4; isCommonPath = $true; Message = "Macro Notification Settings | VBA Macro Notification Settings" },
    @{ Apps = @("Access", "Excel", "Powerpoint", "Publisher", "Visio", "Word"); Path = "security"; Key = "vbadigsigtrustedpublishers"; ExpectedValue = 0; isCommonPath = $true; Message = "Macro Notification Settings | Require Macros to be signed by a trusted publisher" },
    @{ Apps = @("Access", "Excel", "Powerpoint", "Publisher", "Visio", "Word"); Path = "security"; Key = "vbarequirelmtrustedpublisher"; ExpectedValue = 0; isCommonPath = $true; Message = "Macro Notification Settings | Block certificates from trusted publishers that are only installed in the current user certificate store" },
    @{ Apps = @("Access", "Excel", "Powerpoint", "Publisher", "Visio", "Word"); Path = "security"; Key = "vbarequiredigsigwithcodesigningeku"; ExpectedValue = 0; isCommonPath = $true; Message = "Macro Notification Settings | Require Extended Key Usage (EKU) for certificates from trusted publishers" },
    @{ Apps = @("Access", "Excel", "Powerpoint", "Publisher", "Visio", "Word"); Path = "security"; Key = "blockcontentexecutionfrominternet"; ExpectedValue = 1; isCommonPath = $true; Message = "Block macros from running in Office files from the internet" },
    @{ Apps = @("Access", "Excel", "MS Project", "Powerpoint", "Visio", "Word"); Path = "security\Trusted Locations"; Key = "allownetworklocations"; ExpectedValue = 0; isCommonPath = $true; Message = "Allow Trusted Locations on the Network" },
    @{ Apps = @("Access", "Excel", "MS Project", "Powerpoint", "Visio", "Word"); Path = "security\Trusted Locations"; Key = "alllocationsdisabled"; ExpectedValue = 1; isCommonPath = $true;  Message = "Disable all trusted locations" },
    @{ Apps = @("Access", "Excel", "Powerpoint", "Visio", "Word"); Path = "security\Trusted Documents"; Key = "disablenetworktrusteddocuments"; ExpectedValue = 1; isCommonPath = $true;  Message = "Turn off Trusted Documents on the network" },
    @{ Apps = @("Access", "Excel", "Powerpoint", "Visio", "Word"); Path = "security\Trusted Documents"; Key = "disabletrusteddocuments"; ExpectedValue = 1; isCommonPath = $true;  Message = "Turn off trusted documents" },
    @{ Apps = @("Excel", "Powerpoint", "Word"); Path = "security"; Key = "accessvbom"; ExpectedValue = 0; isCommonPath = $true;  Message = "Trust access to Visual Basic Project" },
    @{ Apps = @("Excel"); Path = "security"; Key = "xl4macrowarningfollowvba"; ExpectedValue = 0; isCommonPath = $true;  Message = "Macro Notification Settings | Enable Excel 4.0 macros when VBA macros are enabled" },
    @{ Apps = @("MS Project"); Path = "security"; Key = "vbawarnings"; ExpectedValue = 4; isCommonPath = $true;  Message = "Macro Notification Settings | VBA Macro Notification Settings" },
    @{ Apps = @("Office"); Path = "16.0\common"; Key = "vbaoff"; ExpectedValue = 1; isCommonPath = $false; Message = "Disable VBA for Office applications " },
    @{ Apps = @("Office"); Path = "16.0\common\security\Trusted Locations"; Key = "allow user locations"; ExpectedValue = 0; isCommonPath = $false; Message = "Allow mix of policy and user locations" },
    @{ Apps = @("Office"); Path = "common\security"; Key = "automationsecurity"; ExpectedValue = 3; isCommonPath = $false; Message = "Automation Security | Set the Automation Security Level" }
    @{ Apps = @("Outlook"); Path = "security"; Key = "level"; ExpectedValue = 4; isCommonPath = $true;  Message = "Security setting for macros | Security Level" },
    @{ Apps = @("Outlook"); Path = "security"; Key = "donttrustinstalledfiles"; ExpectedValue = 1; isCommonPath = $true;  Message = "Apply macro security settings to macros, add-ins and additional actions" },
    @{ Apps = @("Publisher"); Path = "common\security"; Key = "automationsecuritypublisher"; ExpectedValue = 3; isCommonPath = $false; Message = "Publisher Automation Security Level | Publisher Automation Security Level" }
    @{ Apps = @("Visio"); Path = "application"; Key = "createvbaprojects"; ExpectedValue = 0; isCommonPath = $true;  Message = "Enable Microsoft Visual Basic for Applications project creation" },
    @{ Apps = @("Visio"); Path = "application"; Key = "loadvbaprojectsfromtext"; ExpectedValue = 0; isCommonPath = $true;  Message = "Load Microsoft Visual Basic for Applications projects from text" }
    # Add other settings here
)

$applications = @("Access", "Excel", "MS Project", "Office", "Outlook", "Powerpoint", "Publisher", "Visio", "Word")

foreach ($app in $applications) {
    foreach ($setting in $settings) {
        if ($app -in $setting.Apps) {
            $$ = $messageFormatted = $setting.Message -f $app
            Test-RegistrySetting -app $app -path $setting.Path -key $setting.Key -expectedValue $setting.ExpectedValue -message $messageFormatted -isCommonPath $setting.isCommonPath
        }
    }
}
