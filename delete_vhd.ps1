        # Azure Storage informasjon
        $storageAccountName = "lukstoracc"          # Sett til ditt storage account navn
        $resourceGroupName = "Lukas-Norway-RG"      # Sett til ditt resource group navn
        $shareName = "fslogix"                      # Sett til ditt fileshare navn
        $user = "joas"

        # Koble til Azure
        Connect-AzAccount

        # Test access to the storage account
        $storageAccount = Get-AzStorageAccount -ResourceGroupName $resourceGroupName -Name $storageAccountName -ErrorAction Stop
        if (-not $storageAccount) {
        throw "Kan ikke få tilgang til lagringskontoen $storageAccountName."
        }


        # Hent Storage Account nøkkel
        $storageKey = (Get-AzStorageAccountKey -ResourceGroupName $resourceGroupName -Name $storageAccountName)[0].Value

        # Opprett en lagringskontekst
        $storageContext = New-AzStorageContext -StorageAccountName $storageAccountName -StorageAccountKey $storageKey

        # Monter fileshare som PSDrive
        # $psDriveName = "FSLogixShare"
        New-PSDrive -Name "FSLogixShare" -PSProvider FileSystem -Root "\\$($storageAccountName).file.core.windows.net\$shareName" -Credential (New-Object PSCredential -ArgumentList "Azure\$storageAccountName", (ConvertTo-SecureString -String $storageKey -AsPlainText -Force))

        # Finn .vhd-filer for brukeren
        $vhdFiles = Get-ChildItem -Path "FSLogixShare:\" -Recurse -Filter "*.vhd*" | Where-Object { $_.Name -like "*$user*" }

        if ($vhdFiles.Count -gt 0) {
            foreach ($vhdFile in $vhdFiles) {
                # Close any open handles to the file
                # Note: Closing open file handles can be risky; ensure it's safe to do so.
                $openFiles = Get-SmbOpenFile | Where-Object { $_.Path -like "*$($vhdFile.FullName)*" }
                foreach ($openFile in $openFiles) {
                    Close-SmbOpenFile -FileId $openFile.FileId -Force
                }

                # Implement a retry mechanism
                $maxRetries = 1
                $retryDelay = 10 # seconds
                $attempt = 0
                $deleted = $false

                while (-not $deleted -and $attempt -lt $maxRetries) {
                    try {
                        Remove-Item -Path $vhdFile.FullName -Force -ErrorAction Stop
                        $deleted = $true
                    } catch {
                        $attempt++
                        Start-Sleep -Seconds $retryDelay
                    }
                }

                if (-not $deleted) {
                    throw "Could not delete file $($vhdFile.FullName) after $maxRetries attempts."
                }
            }
            [System.Windows.Forms.MessageBox]::Show("Brukeren ${name} er offboardet, og tilhørende .vhd-filer er slettet.")
        } else {
            [System.Windows.Forms.MessageBox]::Show("Brukeren ${name} er offboardet, men ingen .vhd-filer ble funnet.")
        }

        # Fjern PSDrive
        Remove-PSDrive -Name "FSLogixShare"