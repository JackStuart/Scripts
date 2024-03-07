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
    @{ Apps = @("Access", "Excel", "MS Project", "Powerpoint", "Publisher", "Visio", "Word"); Path = "security"; Key = "vbawarnings"; ExpectedValue = 4; isCommonPath = $true; Message = "Macro Notification Settings | VBA Macro Notification Settings" },
    @{ Apps = @("Office"); Path = "16.0\common"; Key = "vbaoff"; ExpectedValue = 1; isCommonPath = $false; Message = "Disable VBA for Office applications " },
    @{ Apps = @("Office"); Path = "common\security"; Key = "automationsecurity"; ExpectedValue = 3; isCommonPath = $false; Message = "Automation Security | Set the Automation Security Level" }
    @{ Apps = @("Outlook"); Path = "security"; Key = "level"; ExpectedValue = 4; isCommonPath = $true;  Message = "Security setting for macros | Security Level" },
    @{ Apps = @("Publisher"); Path = "common\security"; Key = "automationsecuritypublisher"; ExpectedValue = 3; isCommonPath = $false; Message = "Publisher Automation Security Level | Publisher Automation Security Level" }
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
