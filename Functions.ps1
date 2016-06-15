############# Windows Azure Storage Functions #############################
Function GetAzureStorageAccount {

Param
    (
    $StorageAccount
    )

    # Test for the storage account
    Try
    {
        # Get the account name supplied, if this failed the catch block will be run.
        Get-AzureStorageAccount -StorageAccountName $StorageAccount -ErrorAction Stop

        # Set the valid variable to true if the command above works ok.
        $Valid = $True 
    }
    Catch
    {
        # A descriptive error!
        Write-Host "Erm...Storage account doesn't exist check for typos" -ForegroundColor Red
    }
    
Return, $Valid
}

Function CheckContainerName {

Param
    (
    $Context,
    $Container
    )

    Try
    {
        Get-AzureStorageContainer -Context $Context -Name $Container -ErrorAction Stop

        $Valid = $true

    }
    Catch
    {
        Write-Host "Erm...container doesn't exist check for typos" -ForegroundColor Red
            
        
    }
Return, $Valid
}

Function CheckVhdName {

Param
    (
    $context,
    $Container,
    $vhd
    )

    Try
    {
        Get-AzureStorageBlob -Context $context -Container $container -Blob $vhd -ErrorAction Stop
        
        $Valid = $true
    }
    Catch
    {
        Write-Host "Erm...blob doesn't exist check for typos" -ForegroundColor Red
        
    }
Return, $Valid
}
############# Windows Azure Storage Functions #############################
############# Exchange Functions ###################################
Function ConnectToExchange {
    
    Param
    (
    $Credentials,
	$Location
    )
    
    # This array is used to return multiple values back to the script.
    [array]$arrExchConnection = @()
	
	Try
    {
        # Check the Exchange PowerShell location
		If ($Location -eq "L") {
			
			$ConnectionUri = Read-Host "Please enter the full Exchange PowerShell Url"		
		
		}
		Else {
		
			$ConnectionUri = "https://ps.outlook.com/PowerShell"
		
		}
		
		# Attempt to connect to Exchange Online.
        $arrExchConnection = New-PSSession -ConfigurationName Microsoft.Exchange `
        -ConnectionUri $ConnectionUri -Credential $Credentials `
        -AllowRedirection -Authentication Basic -ErrorAction Stop

        # If the connection above worked then add value to the array. If the connection above fails then the -ErrorAction Stop will 
        # jump straight to the catch block.
        $arrExchConnection += $true
    }
    Catch
    {
        # If the connection fails then warnings below are outputted.
        Write-Host "Invalid credentials, please re-enter your credentials." -ForegroundColor Yellow
		# The place holder is required to fill the first element in the array 
		# as I check for the second element to validate whether the connection 
		# was succeeded or failed. 
        $arrExchConnection = "Place Holder"
		$arrExchConnection += $false
    }
Return, $arrExchConnection
}
############# Exchange Online Functions ###################################
############# Office 365 Functions ########################################
Function ConnectToMicrosoftOnline {

	Param
	(
	$Credentials
	)
	
	Try 
	{
		Connect-MsolService -Credential $credentials -ErrorAction Stop
		$ConnectToMicrosoftOnline = $true
	}
	Catch
	{
		Write-Host "Invalid credentials. Please re-enter" -ForegroundColor Red
		$ConnectToMicrosoftOnline = $false
	}
Return, $ConnectToMicrosoftOnline
}
############# Office 365 Functions ########################################
############# Windows Azure Connectivity Functions ########################
Function ConnnectToWindowsAzure {

Param
    (
    $credentials
    )


}
############# Windows Azure Connectivity Functions ########################
############# General Regex Functions #####################################
Function CheckDomain {
	
	Param
    (
    $sDomainName
    )
	
		if ($sDomainName -notmatch "^(?!www\.)(?!\.)([a-z0-9\-\.]+)+((com)|(me\.uk)|(co\.uk)|(eu)|(org)|(ac))$"){
            Write-host $sDomainName "does not match the expected format for a domain name e.g. site.com|.co.uk|.eu|.org | .ac is valid" -ForegroundColor Red
            $Validator = $false
        }
		else {
	
			$Validator = $true	
		}	
	Return $Validator
}
############# General Regex Functions #####################################
############# General Functions #####################################
Function WriteToLog {

	Param
	(
	$sLogFile,
	$sLogContent
	)
	
	If (Test-Path $sLogFile) {
	
		# Append to the file
		Add-Content -Path $sLogFile -Value $sLogContent
		
	}
	Else {
	
		# Create the file
		New-Item -ItemType File -Path $sLogFile
		Add-Content -Path $sLogFile -Value $sLogContent
		
	}
}

Function ReadReg
{
    $temp = $Args[0]
    $temp1 = $Args[1]
    $Val = get-itemproperty -path $temp
    Return $Val.$Temp1
}

Function CheckDirectory {

    Param
    (
    $sFolderPath
    )

    If (-not (Test-Path $sFolderPath)) {

        # Create the folder in the PSScriptRoot
        New-Item -ItemType Directory $sFolderPath

    }

}


Function Load_ActiveDirectoryTools
{
    # Read the registry to obtain the operating system version.
	$sVer = ReadReg "Hklm:\Software\Microsoft\Windows NT\CurrentVersion\" "CurrentVersion"
		
		switch ($sVer)
		{
			Default
			{
				Write-Host "Unsupported operating system version" -ForegroundColor Red
			}
			"6.0"
			{
				while (!$Valid)
				{
					Add-PSSnapin WebAdministration
						if (!$?) {
				  			Write-Host "Windows 2008 detected, but no PS Snapin Available - Download from: http://www.iis.net/download/PowerShell" -ForegroundColor Red
               				Read-Host "Please install the WebAdministraton SnapIn, then press enter"
            			}
            			else{
							$Valid = $true
						}
				}
			}
			"6.1"
			{
				while (!$Valid)
				{
					Import-Module WebAdministration
						if (!$?) {
                			Write-Host "Windows R2 detected, but no WebAdministration Module Available - Install via Server Manager" -ForegroundColor Red
                			Read-Host "Please install the WebAdministraton module, then press enter"
            			}
						else{
            				$Valid = $true
						}
				}	
			}
		}	
Return $Valid
}
############# General Functions #####################################