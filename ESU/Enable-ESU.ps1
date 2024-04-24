$subscriptionName = "" #Enter the name of the subscription
$rgNameArcManagement = "" #Enter the name of the Resource Group you have the Arc resources/liucences in 
$region = "australiaeast" #Region

#array of Servers
$servers = @(
  ""
    )

    foreach ($serverName in $servers) {
        $esuLicenseName = "ESU-$($ServerName)" #ESU Licence Name
        ##### START OF PROVISIONING LICENCE #####
        $arcMachine = Get-AzConnectedMachine -Name $serverName -ResourceGroupName $rgNameArcManagement
    
        #Set's the minimum core count
        [int]$coreCount = ($arcMachine.DetectedProperty | ConvertFrom-Json).corecount
        if ($coreCount -lt 8) {
            $vCores = 8
        }
        else {
            $vCores = $coreCount
        }
    
    
        $accessToken = (Get-AzAccessToken -ResourceUrl https://management.azure.com).Token
        $subName = Get-AzSubscription | Where-Object { $_.Name -like $subscriptionName } -ErrorAction SilentlyContinue
        $esuID = $subName.SubscriptionId + "/resourceGroups/" + $rgNameArcManagement + "/providers/Microsoft.HybridCompute/licenses/" + $esuLicenseName
    
        $URI = "https://management.azure.com/subscriptions/" + $esuID + "?api-version=2023-06-20-preview"
        $headers = [ordered]@{"Content-Type" = "application/json"; "Authorization" = "Bearer $accessToken" }
        $method = "PUT"
        $jsonObject = @{
            location   = $region
            properties = @{
                licenseDetails = @{
                    state      = "Activated"
                    target     = "Windows Server 2012"
                    Edition    = "Standard"
                    Type       = "vCore"
                    Processors = $vCores
                }
            }
        }
        $jsonBody = $jsonObject | ConvertTo-Json
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 
     
        # Send the HTTP request to the specified URI with the provided method, headers, and body
        $response = Invoke-WebRequest -URI $URI -Method $method -Headers $headers -Body $jsonBody -TimeoutSec 90
    
        Write-Output $response
        ##### END OF PROVISIONING LICENCE #####
    
    
    
    
        ##### START OF ASSIGNING LICENCE #####
        Start-Sleep -seconds 10
    
        $URI = "https://management.azure.com" + $arcMachine.Id + "/licenseProfiles/default?api-version=2023-06-20-preview"
        $headers = [ordered]@{"Content-Type" = "application/json"; "Authorization" = "Bearer $accessToken" }
        $method = "PUT"
        $jsonObject = @{
            location   = $region
            properties = @{
                esuProfile = @{
                    assignedLicense = "/subscriptions/$($esuID)"
                }
            }
        }
        $jsonBody = $jsonObject | ConvertTo-Json
        $response = Invoke-WebRequest -URI $URI -Method $method -Headers $headers -Body $jsonBody -TimeoutSec 90
    }
    
    ##### END OF ASSIGNING LICENCE #####
