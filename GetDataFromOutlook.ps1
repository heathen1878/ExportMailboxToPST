#############################################################################################
#                                                                                           #                            
#                                                                                           #
#############################################################################################

# Get the location where the script executes from. This will be where the logs and pst files are output to.
$StartIn = $PSScriptRoot

# Set the Error Action Preference to stop. Error will be handled by the Try Catch Finally blocks
$Global:ErrorActionPreference="Stop"

. (Join-Path -Path $StartIn -ChildPath "\functions.ps1")

# Start of script.
WriteToLog -sLogFile (Join-Path -path $StartIn -ChildPath "getDataFromOutlook.log") -sLogContent "Script started...$(Get-Date)"

# The pst files will be exported to an exports folder on that users desktop.
CheckDirectory -sFolderPath (Join-Path -path $StartIn -ChildPath "\Exports")

# Create an Outlook object
$objOutlook = New-object -ComObject Outlook.Application

# Output Outlook version to the log
WriteToLog -sLogFile (Join-Path -path $StartIn -ChildPath "getDataFromOutlook.log") -sLogContent "Outlook version...$($objOutlook.Version)"

# Get the MAPI namespaces associated with your Outlook profile. If you need to export multiple users then these
# users should be included in your profile. The mailbox mappings should also be online in order to get all emails.
$objNamespaces = $objOutlook.getNamespace("MAPI")

# Output number of stores being processed
WriteToLog -sLogFile (Join-Path -path $StartIn -ChildPath "getDataFromOutlook.log") -sLogContent "No. of stores found...$($objNamespaces.Stores.Count)"

# Loop through all the stores within the profile, create a personal file, and then export all items to that file. 
# Once the export is complete, disconnect the personal file. 
$objNamespaces.Stores | % {

   $ProfileName = $_.DisplayName

   WriteToLog -sLogFile (Join-Path -path $StartIn -ChildPath "getDataFromOutlook.log") -sLogContent "Working on: $($ProfileName)"

   # Check if the store is cached or online.
   If ($_.IsCachedExchange -eq "True") {
        
        # If Exchange is cached output to a log that the mailbox content is probably not complete.
        # Registry reference - http://www.bytemedev.com/powershell-disable-cached-mode-in-outlook-20102013/
        WriteToLog -sLogFile (Join-Path -path $StartIn -ChildPath "getDataFromOutlook.log") -sLogContent "Found mail profile which is cached...$ProfileName"
        WriteToLog -sLogFile (Join-Path -path $StartIn -ChildPath "getDataFromOutlook.log") -sLogContent "Modify the registry on the local computer to disable Outlook cached mode. Use the following PowerShell"
        WriteToLog -sLogFile (Join-Path -path $StartIn -ChildPath "getDataFromOutlook.log") -sLogContent 'New-Item -Path "HKCU:\SOFTWARE\Policies\Microsoft\Office\15.0\Outlook\Cached Mode" -Force'
        WriteToLog -sLogFile (Join-Path -path $StartIn -ChildPath "getDataFromOutlook.log") -sLogContent 'New-ItemProperty -Path "HKCU:\SOFTWARE\Policies\Microsoft\Office\15.0\Outlook\Cached Mode" -Name "Enable" -Value 0 -PropertyType "DWORD" -Force'
        
   }

   # Define the pst file where data will be stored. 
   $PstExportFile = -join ((Join-Path -path $StartIn -ChildPath "\Exports\"),$ProfileName,".pst")

   # Add the pst file to Outlook.
   $objNamespaces.AddStore($PstExportFile)
   
   # Get the Outlook store which is the pst file you have just added.
   $PstStore = $objNamespaces.Stores | ? {$_.FilePath -eq $PstExportFile}

   # Name the store which is the pst file. This will be used later on to dismount the store.
   # Reference - https://social.technet.microsoft.com/Forums/scriptcenter/en-US/c76c7167-8336-4261-ac40-2fb44ff3b3f3/powershell-and-outlook-removestore-method?forum=ITCG
   $PstStoreName = $PstStore.GetRootFolder()
   $PstStoreName.Name = "PST: $ProfileName"

   # Loop through all folders within the store
   $_.GetRootFolder().Folders | % {

        Try {
        
            WriteToLog -sLogFile (Join-Path -path $StartIn -ChildPath "getDataFromOutlook.log") -sLogContent "$($_.Name)"
            $_.CopyTo($PstStore)

        }
        Catch {

            WriteToLog -sLogFile (Join-Path -path $StartIn -ChildPath "getDataFromOutlook.log") -sLogContent "$($Error[0].Exception.Message)"

        }
   }

   # Disconnect the pst store and move onto the next mail store.
   $objNamespaces.RemoveStore($PstStoreName)
}