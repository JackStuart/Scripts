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
        Write-Output $message
    }
}

$settings = @(
    @{ Apps = @("Access", "Excel", "Powerpoint", "Publisher", "Visio", "Word"); Path = "security"; Key = "vbawarnings"; ExpectedValue = 4; isCommonPath = $true; Message = "Macro Notification Settings - Disabled without notification failed for {0}" },
    @{ Apps = @("Access", "Excel", "Powerpoint", "Publisher", "Visio", "Word"); Path = "security"; Key = "vbadigsigtrustedpublishers"; ExpectedValue = 0; isCommonPath = $true; Message = "Macro Notification Settings - Require Macros to be signed by a trusted publisher failed for {0}" },
    @{ Apps = @("Access", "Excel", "Powerpoint", "Publisher", "Visio", "Word"); Path = "security"; Key = "vbarequirelmtrustedpublisher"; ExpectedValue = 0; isCommonPath = $true; Message = "Macro Notification Settings - Block certificates from trusted publishers that are only installed in the current user certificate store failed for {0}" },
    @{ Apps = @("Access", "Excel", "Powerpoint", "Publisher", "Visio", "Word"); Path = "security"; Key = "vbarequiredigsigwithcodesigningeku"; ExpectedValue = 0; isCommonPath = $true; Message = "Macro Notification Settings - Require Extended Key Usage (EKU) for certificates from trusted publishers failed for {0}" },
    @{ Apps = @("Access", "Excel", "Powerpoint", "Publisher", "Visio", "Word"); Path = "security"; Key = "blockcontentexecutionfrominternet"; ExpectedValue = 1; isCommonPath = $true; Message = "Block macros from running in Office files from the internet failed for {0}" },
    @{ Apps = @("Access", "Excel", "MS Project", "Powerpoint", "Visio", "Word"); Path = "security\Trusted Locations"; Key = "allownetworklocations"; ExpectedValue = 0; isCommonPath = $true; Message = "Disable Allow Trusted Locations on the Network failed for {0}" },
    @{ Apps = @("Access", "Excel", "MS Project", "Powerpoint", "Visio", "Word"); Path = "security\Trusted Locations"; Key = "alllocationsdisabled"; ExpectedValue = 1; isCommonPath = $true;  Message = "Disable all Trusted Location failed for {0}" },
    @{ Apps = @("Access", "Excel", "Powerpoint", "Visio", "Word"); Path = "security\Trusted Documents"; Key = "disablenetworktrusteddocuments"; ExpectedValue = 1; isCommonPath = $true;  Message = "Turn off Trusted Documents on the network failed for {0}" },
    @{ Apps = @("Access", "Excel", "Powerpoint", "Visio", "Word"); Path = "security\Trusted Documents"; Key = "disabletrusteddocuments"; ExpectedValue = 1; isCommonPath = $true;  Message = "Turn off trusted documents failed for {0}" },
    @{ Apps = @("Excel", "Powerpoint", "Word"); Path = "security"; Key = "accessvbom"; ExpectedValue = 0; isCommonPath = $true;  Message = "Test for Trust access to Visual Basic Project failed for {0}" },
    @{ Apps = @("Excel"); Path = "security"; Key = "xl4macrowarningfollowvba"; ExpectedValue = 0; isCommonPath = $true;  Message = "Enable Excel 4.0 macros when VBA macros are enabled failed for {0}" },
    @{ Apps = @("Office"); Path = "16.0\common"; Key = "vbaoff"; ExpectedValue = 1; isCommonPath = $false; Message = "Disable VBA for Office applications failed for {0}" },
    @{ Apps = @("Office"); Path = "16.0\common\security\Trusted Locations"; Key = "allow user locations"; ExpectedValue = 0; isCommonPath = $false; Message = "Allow mix of policy and user locations failed for {0}" },
    @{ Apps = @("Office"); Path = "common\security"; Key = "automationsecurity"; ExpectedValue = 3; isCommonPath = $false; Message = "Automation Security failed for {0}" }
    @{ Apps = @("Outlook"); Path = "security"; Key = "level"; ExpectedValue = 4; isCommonPath = $true;  Message = "Security setting for macros failed for {0}" },
    @{ Apps = @("Outlook"); Path = "security"; Key = "donttrustinstalledfiles"; ExpectedValue = 1; isCommonPath = $true;  Message = "Apply macro security settings to macros, add-ins and additional actions failed for {0}" },
    @{ Apps = @("MS Project"); Path = "security"; Key = "vbawarnings"; ExpectedValue = 4; isCommonPath = $true;  Message = "Macro Notification Settings failed for {0}" },
    @{ Apps = @("Publisher"); Path = "common\security"; Key = "automationsecuritypublisher"; ExpectedValue = 3; isCommonPath = $false; Message = "Publisher Automation Security failed for {0}" }
    @{ Apps = @("Visio"); Path = "application"; Key = "createvbaprojects"; ExpectedValue = 0; isCommonPath = $true;  Message = "Enable Microsoft Visual Basic for Applications project creation failed for {0}" },
    @{ Apps = @("Visio"); Path = "application"; Key = "loadvbaprojectsfromtext"; ExpectedValue = 0; isCommonPath = $true;  Message = "Load Microsoft Visual Basic for Applications projects from text failed for {0}" }
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
