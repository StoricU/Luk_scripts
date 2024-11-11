# Importer nødvendige moduler
Import-Module ActiveDirectory
Import-Module Az.Accounts
Import-Module Az.Storage

# Last inn nødvendige assemblies for GUI
[void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
[void][System.Reflection.Assembly]::LoadWithPartialName('System.Drawing')

# Sett testmodus (sett til $false for å utføre faktiske endringer)
$testMode = $true   # Sett til $false for å kjøre faktisk sletting

# Opprett form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Offboard Bruker - Lukas Stiftelsen"
$form.Size = New-Object System.Drawing.Size(300,150)
$form.StartPosition = "CenterScreen"

# Opprett label for brukernavn
$userLabel = New-Object System.Windows.Forms.Label
$userLabel.Location = New-Object System.Drawing.Size(10,20)
$userLabel.Size = New-Object System.Drawing.Size(80,20)
$userLabel.Text = "Brukernavn:"
$form.Controls.Add($userLabel)

# Opprett tekstboks for brukernavn
$userTextBox = New-Object System.Windows.Forms.TextBox
$userTextBox.Location = New-Object System.Drawing.Size(100,20)
$userTextBox.Size = New-Object System.Drawing.Size(180,20)
$form.Controls.Add($userTextBox)

# Opprett knapp for offboarding
$offboardButton = New-Object System.Windows.Forms.Button
$offboardButton.Location = New-Object System.Drawing.Size(100,60)
$offboardButton.Size = New-Object System.Drawing.Size(100,30)
$offboardButton.Text = "Offboard"
$form.Controls.Add($offboardButton)

# Legg til klikk-hendelse for knappen
$offboardButton.Add_Click({
    # Hent brukernavn fra tekstboksen
    $username = $userTextBox.Text.Trim()

    # Sjekk om brukernavn er fylt ut
    if ([string]::IsNullOrWhiteSpace($username)) {
        [System.Windows.Forms.MessageBox]::Show("Vennligst oppgi brukernavn.")
        return
    }

    # Forsøk å hente brukeren fra AD
    try {
        $user = Get-ADUser -Identity $username -Properties *
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Fant ikke brukeren $username i Active Directory.")
        return
    }

    try {
        if (-not $testMode) {
            # Deaktiver brukerkontoen
            Disable-ADAccount -Identity $user

            # Flytt brukeren til "Disabled Users" OU
            $disabledUsersOU = "OU=Disabled Users,DC=lukasstiftelsen,DC=no"
            Move-ADObject -Identity $user.DistinguishedName -TargetPath $disabledUsersOU

            # Fjern brukeren fra alle grupper unntatt "Domain Users"
            $groups = Get-ADPrincipalGroupMembership -Identity $user | Where-Object { $_.Name -ne 'Domain Users' }
            foreach ($group in $groups) {
                Remove-ADGroupMember -Identity $group -Members $user -Confirm:$false
            }
        } else {
            # Simulerer deaktivering av brukerkontoen
            Write-Host "Simulerer deaktivering av brukerkontoen $username..."

            # Simulerer flytting av brukeren til "Disabled Users" OU
            Write-Host "Simulerer flytting av brukeren til OU=Disabled Users,DC=lukasstiftelsen,DC=no..."

            # Simulerer fjerning av brukeren fra alle grupper unntatt 'Domain Users'
            Write-Host "Simulerer fjerning av brukeren fra alle grupper unntatt 'Domain Users'..."
        }

        # Azure Storage informasjon
        $storageAccountName = "lukstoracc"          # Sett til ditt storage account navn
        $resourceGroupName = "Lukas-Norway-RG"      # Sett til ditt resource group navn
        $shareName = "fslogix"                      # Sett til ditt fileshare navn

        if (-not $testMode) {
            # Koble til Azure hvis ikke allerede tilkoblet
            if (-not (Get-AzContext)) {
                Connect-AzAccount
            }

            # Hent Storage Account nøkkel
            $storageKey = (Get-AzStorageAccountKey -ResourceGroupName $resourceGroupName -Name $storageAccountName)[0].Value

            # Opprett en lagringskontekst
            $storageContext = New-AzStorageContext -StorageAccountName $storageAccountName -StorageAccountKey $storageKey

            # Monter fileshare som PSDrive
            $psDriveName = "FSLogixShare"
            New-PSDrive -Name $psDriveName -PSProvider FileSystem -Root "\\$($storageAccountName).file.core.windows.net\$shareName" -Credential (New-Object PSCredential -ArgumentList "Azure\$storageAccountName", (ConvertTo-SecureString -String $storageKey -AsPlainText -Force)) -ErrorAction Stop

            # Finn .vhd-filer for brukeren
            $vhdFileNamePattern = "$username*.vhd*"

            # Hent filene
            $vhdFiles = Get-ChildItem -Path "$psDriveName:\" -Recurse -Filter $vhdFileNamePattern

            if ($vhdFiles.Count -gt 0) {
                # Slett .vhd-filene
                foreach ($vhdFile in $vhdFiles) {
                    Remove-Item -Path $vhdFile.FullName -Force
                }
                [System.Windows.Forms.MessageBox]::Show("Brukeren $username er offboardet, og tilhørende .vhd-filer er slettet.")
            } else {
                [System.Windows.Forms.MessageBox]::Show("Brukeren $username er offboardet, men ingen .vhd-filer ble funnet.")
            }

            # Fjern PSDrive
            Remove-PSDrive -Name $psDriveName
        } else {
            # Simulerer tilkobling til Azure
            Write-Host "Simulerer tilkobling til Azure..."

            # Simulerer henting av Storage Account nøkkel
            Write-Host "Simulerer henting av Storage Account nøkkel for $storageAccountName..."

            # Simulerer opprettelse av lagringskontekst
            Write-Host "Simulerer opprettelse av lagringskontekst..."

            # Simulerer montering av fileshare som PSDrive
            Write-Host "Simulerer montering av fileshare som PSDrive..."

            # Simulerer søk etter .vhd-filer for brukeren
            $vhdFileNamePattern = "$username*.vhd*"
            Write-Host "Simulerer søk etter .vhd-filer med mønster '$vhdFileNamePattern'..."

            # Simulerer funn av VHD-filer (for testformål)
            $vhdFiles = @("FSLogixShare\$username_profile.vhd", "FSLogixShare\$username_data.vhd")

            if ($vhdFiles.Count -gt 0) {
                # Simulerer sletting av .vhd-filene
                foreach ($vhdFile in $vhdFiles) {
                    Write-Host "Simulerer sletting av fil: $vhdFile"
                }
                [System.Windows.Forms.MessageBox]::Show("Test fullført: Brukeren $username ville blitt offboardet, og tilhørende .vhd-filer ville blitt slettet.")
            } else {
                [System.Windows.Forms.MessageBox]::Show("Test fullført: Brukeren $username ville blitt offboardet, men ingen .vhd-filer ble funnet.")
            }

            # Simulerer fjerning av PSDrive
            Write-Host "Simulerer fjerning av PSDrive..."
        }

        # Lukk formen
        $form.Close()
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Feil ved offboarding av brukeren $username: $($_.Exception.Message)")
    }
})

# Vis formen
$form.ShowDialog()
