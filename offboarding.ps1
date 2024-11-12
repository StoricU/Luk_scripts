# Importer nødvendige moduler
Import-Module ActiveDirectory
Import-Module Az.Accounts
Import-Module Az.Storage

# Last inn nødvendige assemblies for GUI
[void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
[void][System.Reflection.Assembly]::LoadWithPartialName('System.Drawing')

# Funksjon for å generere potensielle brukernavn basert på navnet
function Generate-PossibleUsernames {
    param (
        [string]$firstName,
        [string]$lastName
    )

    # Fjern spesialtegn og gjør om til små bokstaver
    $firstNameClean = $firstName.Replace("ø", "o").Replace("æ", "ae").Replace("å", "a").Replace(" ", "").ToLower()
    $lastNameClean = $lastName.Replace("ø", "o").Replace("æ", "ae").Replace("å", "a").Replace(" ", "").ToLower()

    $usernameList = @()

    # Forsøk 1: Første 2 bokstaver av fornavn + første 2 bokstaver av etternavn
    if ($firstNameClean.Length -ge 2 -and $lastNameClean.Length -ge 2) {
        $username1 = $firstNameClean.Substring(0,2) + $lastNameClean.Substring(0,2)
        $usernameList += $username1
    }

    # Forsøk 2: Første 2 bokstaver av fornavn + første 3 bokstaver av etternavn
    if ($firstNameClean.Length -ge 2 -and $lastNameClean.Length -ge 3) {
        $username2 = $firstNameClean.Substring(0,2) + $lastNameClean.Substring(0,3)
        $usernameList += $username2
    }

    return $usernameList
}

# Opprett form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Offboard Bruker - Lukas Stiftelsen"
$form.Size = New-Object System.Drawing.Size(300,150)
$form.StartPosition = "CenterScreen"

# Opprett label for navn
$nameLabel = New-Object System.Windows.Forms.Label
$nameLabel.Location = New-Object System.Drawing.Size(10,20)
$nameLabel.Size = New-Object System.Drawing.Size(80,20)
$nameLabel.Text = "Navn:"
$form.Controls.Add($nameLabel)

# Opprett tekstboks for navn
$nameTextBox = New-Object System.Windows.Forms.TextBox
$nameTextBox.Location = New-Object System.Drawing.Size(100,20)
$nameTextBox.Size = New-Object System.Drawing.Size(180,20)
$form.Controls.Add($nameTextBox)

# Opprett knapp for offboarding
$offboardButton = New-Object System.Windows.Forms.Button
$offboardButton.Location = New-Object System.Drawing.Size(100,60)
$offboardButton.Size = New-Object System.Drawing.Size(100,30)
$offboardButton.Text = "Offboard"
$form.Controls.Add($offboardButton)

# Legg til klikk-hendelse for knappen
$offboardButton.Add_Click({
    # Hent navn fra formen
    $name = $nameTextBox.Text.Trim()

    # Sjekk om navn er fylt ut
    if ([string]::IsNullOrWhiteSpace($name)) {
        [System.Windows.Forms.MessageBox]::Show("Vennligst fyll ut navnet.")
        return
    }

    # Splitt navn til array av navn
    $names = $name -split " "

    # Tildel fornavn (første ord)
    $firstName = $names[0]

    # Sjekk om minst fornavn og etternavn er oppgitt
    if ($names.Length -ge 2) {
        $lastName = $names[$names.Length - 1]
    } else {
        # Hvis kun ett navn er oppgitt
        [System.Windows.Forms.MessageBox]::Show("Vennligst oppgi både fornavn og etternavn.")
        return
    }

    # Generer mulige brukernavn basert på samme logikk som opprettingsscriptet
    $possibleUsernames = Generate-PossibleUsernames -firstName $firstName -lastName $lastName

    # Finn brukeren i AD basert på mulige brukernavn
    $user = $null
    foreach ($username in $possibleUsernames) {
        $user = Get-ADUser -Filter { SamAccountName -eq $username } -Properties *
        if ($user) {
            break
        }
    }

    if (-not $user) {
        [System.Windows.Forms.MessageBox]::Show("Kan ikke finne brukeren $name i Active Directory. Vennligst sjekk at navnet er korrekt eller kontakt administrator.")
        return
    }

    # Fortsett med offboarding-prosessen
    try {
        # Deaktiver brukerkontoen
        Disable-ADAccount -Identity $user -ErrorAction Stop

        # Output user's distinguished name for verification
        Write-Host "Brukerens DN: $($user.DistinguishedName)"

        # Fjern brukeren fra alle grupper unntatt "Domain Users"
        $groups = Get-ADPrincipalGroupMembership -Identity $user -ErrorAction Stop
        foreach ($group in $groups) {
            if ($group.Name -ne "Domain Users") {
                Remove-ADGroupMember -Identity $group -Members $user -Confirm:$false -ErrorAction Stop
            }
        }

        # Finn "Deaktiverte brukere" OU dynamisk
        $disabledUsersOU = Get-ADOrganizationalUnit -Filter { Name -eq "Deaktiverte brukere" }

        if ($disabledUsersOU) {
            # Flytt brukeren til "Deaktiverte brukere" OU
            Move-ADObject -Identity $user.DistinguishedName -TargetPath $disabledUsersOU.DistinguishedName -ErrorAction Stop
        } else {
            throw "Mål-OU 'Deaktiverte brukere' ble ikke funnet i domenet."
        }

        # Azure Storage informasjon
        $storageAccountName = "lukstoracc"          # Sett til ditt storage account navn
        $resourceGroupName = "Lukas-Norway-RG"      # Sett til ditt resource group navn
        $shareName = "fslogix"                      # Sett til ditt fileshare navn

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
        $psDriveName = "FSLogixShare"
        New-PSDrive -Name $psDriveName -PSProvider FileSystem -Root "\\$($storageAccountName).file.core.windows.net\$shareName" -Credential (New-Object PSCredential -ArgumentList "Azure\$storageAccountName", (ConvertTo-SecureString -String $storageKey -AsPlainText -Force)) -ErrorAction Stop

        # Finn .vhd-filer for brukeren
        $vhdFiles = Get-ChildItem -Path "${psDriveName}:\\" -Recurse -Filter "*.vhd*" | Where-Object { $_.Name -like "*$($user.SamAccountName)*" }

        if ($vhdFiles.Count -gt 0) {
            foreach ($vhdFile in $vhdFiles) {
                # Close any open handles to the file
                # Note: Closing open file handles can be risky; ensure it's safe to do so.
                $openFiles = Get-SmbOpenFile | Where-Object { $_.Path -like "*$($vhdFile.FullName)*" }
                foreach ($openFile in $openFiles) {
                    Close-SmbOpenFile -FileId $openFile.FileId -Force
                }

                # Implement a retry mechanism
                $maxRetries = 5
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
        Remove-PSDrive -Name $psDriveName

        # Lukk formen
        $form.Close()
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Feil ved offboarding av brukeren ${name}: $($_.Exception.Message)")
    }
})

# Vis formen
$form.ShowDialog()
