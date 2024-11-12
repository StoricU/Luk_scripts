# Importer nødvendige moduler
Import-Module ActiveDirectory

# Variabler
$storageAccountName = "lukstoracc"
$resourceGroupName = "Lukas-Norway-RG"
$shareName = "fslogix"
$user = "joas" # Legg til brukernavnet her

# Koble til Azure
Connect-AzAccount

# Hent Storage Account nøkkel
$storageKey = (Get-AzStorageAccountKey -ResourceGroupName $resourceGroupName -Name $storageAccountName)[0].Value

# Opprett en lagringskontekst
$storageContext = New-AzStorageContext -StorageAccountName $storageAccountName -StorageAccountKey $storageKey

# Monter fileshare som PSDrive
New-PSDrive -Name "FSLogixShare" -PSProvider FileSystem -Root "\\$($storageAccountName).file.core.windows.net\$shareName" -Credential (New-Object PSCredential -ArgumentList "Azure\$storageAccountName", (ConvertTo-SecureString -String $storageKey -AsPlainText -Force))

# Finn alle .vhd-filer i fileshare for brukeren "joas"
Get-ChildItem -Path "FSLogixShare:\" -Recurse -Filter "*.vhd" | Where-Object { $_.Name -like "*$user*" } | ForEach-Object {
    Write-Host ".vhd-fil for $user funnet: $($_.FullName) - Sist endret: $($_.LastWriteTime)"
}

# Fjern PSDrive etter ferdig bruk
Remove-PSDrive -Name "FSLogixShare"
