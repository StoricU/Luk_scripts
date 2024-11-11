[void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')

# Opprett form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Ny AD-bruker - Lukas Stiftelsen"
$form.Size = New-Object System.Drawing.Size(300,200)
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

# Opprett label for avdeling
$departmentLabel = New-Object System.Windows.Forms.Label
$departmentLabel.Location = New-Object System.Drawing.Size(10,50)
$departmentLabel.Size = New-Object System.Drawing.Size(80,20)
$departmentLabel.Text = "Avdeling:"
$form.Controls.Add($departmentLabel)

# Opprett komboboks for avdeling
$departmentComboBox = New-Object System.Windows.Forms.ComboBox
$departmentComboBox.Location = New-Object System.Drawing.Size(100,50)
$departmentComboBox.Size = New-Object System.Drawing.Size(180,20)
$departmentComboBox.DropDownStyle = "DropDownList"
$departmentComboBox.Items.Add("Skjelfoss")
$departmentComboBox.Items.Add("Lukasstiftelsen")
$departmentComboBox.Items.Add("Lukas Hospice")
$departmentComboBox.Items.Add("Betania Malvik")
$departmentComboBox.Items.Add("Avlastningsenheten (Betania Malvik)")
$departmentComboBox.Items.Add("Betania Sparbu")
$form.Controls.Add($departmentComboBox)

# Opprett knapp for å opprette bruker
$createButton = New-Object System.Windows.Forms.Button
$createButton.Location = New-Object System.Drawing.Size(100,80)
$createButton.Size = New-Object System.Drawing.Size(100,20)
$createButton.Text = "Opprett"
$createButton.Add_Click({
    # Hent navn og avdeling fra formen
    $name = $nameTextBox.Text
    $department = $departmentComboBox.SelectedItem

    # Sjekk om navn og avdeling er fylt ut
    if ([string]::IsNullOrWhiteSpace($name) -or [string]::IsNullOrWhiteSpace($department)) {
        [System.Windows.Forms.MessageBox]::Show("Vennligst fyll ut både navn og avdeling.")
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

    # Funksjon for å generere unikt brukernavn
    function Generate-UniqueUsername {
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

        # Prøv hvert brukernavn til et unikt er funnet
        foreach ($username in $usernameList) {
            if (-not (Get-ADUser -Filter {SamAccountName -eq $username})) {
                return $username
            }
        }

        # Hvis ingen unike brukernavn er funnet, returner null
        return $null
    }

    # Generer unikt brukernavn
    $username = Generate-UniqueUsername -firstName $firstName -lastName $lastName

    if ($null -eq $username) {
        [System.Windows.Forms.MessageBox]::Show("Kan ikke generere et unikt brukernavn for $name ved å bruke de gitte reglene. Vennligst prøv med et annet navn eller kontakt administrator.")
        return
    }

    # Generer e-postdomene basert på avdeling
    switch ($department) {
        "Skjelfoss" {
            $emailDomain = "skjelfoss.no"
        }
        "Lukasstiftelsen" {
            $emailDomain = "lukasstiftelsen.no"
        }
        "Betania Malvik" {
            $emailDomain = "betaniamalvik.no"
        }
        "Avlastningsenheten (Betania Malvik)" {
            $emailDomain = "betaniamalvik.no"
        }
        "Lukas Hospice" {
            $emailDomain = "betaniamalvik.no"
        }
        "Betania Sparbu" {
            $emailDomain = "betaniasparbu.no"
        }
        default {
            $emailDomain = "lukasstiftelsen.no"
        }
    }

    # Generer e-post
    $email = "$username@$emailDomain"

    # Sett avdeling og proxy-adresser basert på avdeling
    if ($department -eq "Skjelfoss") {
        $template = Get-ADUser -Identity malskj
        $ou = "OU=Skjelfoss,OU=TS-Users,DC=lukasstiftelsen,DC=no"
        $proxyAddresses = @("smtp:$firstName.$lastName@$emailDomain","SMTP:$username@$emailDomain")
    }
    elseif ($department -eq "Lukasstiftelsen") {
        $template = Get-ADUser -Identity malls
        $ou = "OU=Lukasstiftelsen,OU=TS-Users,DC=lukasstiftelsen,DC=no"
        $proxyAddresses = @("smtp:$firstName.$lastName@$emailDomain")
    }
    elseif ($department -eq "Betania Malvik") {
        $template = Get-ADUser -Identity malbm
        $ou = "OU=Betania Malvik,OU=TS-Users,DC=lukasstiftelsen,DC=no"
        $proxyAddresses = @("smtp:$firstName.$lastName@$emailDomain","SMTP:$username@$emailDomain")
    }
    elseif ($department -eq "Avlastningsenheten (Betania Malvik)") {
        $template = Get-ADUser -Identity malavl
        $ou = "OU=Betania Malvik,OU=TS-Users,DC=lukasstiftelsen,DC=no"
        $proxyAddresses = @("smtp:$firstName.$lastName@$emailDomain","SMTP:$username@$emailDomain")
    }
    elseif ($department -eq "Lukas Hospice") {
        $template = Get-ADUser -Identity mallukhospice
        $ou = "OU=Betania Malvik,OU=TS-Users,DC=lukasstiftelsen,DC=no"
        $proxyAddresses = @("smtp:$firstName.$lastName@$emailDomain","SMTP:$username@$emailDomain")
    }
    elseif ($department -eq "Betania Sparbu") {
        $template = Get-ADUser -Identity malbs
        $ou = "OU=Betania Sparbu,OU=TS-Users,DC=lukasstiftelsen,DC=no"
        $proxyAddresses = @("smtp:$firstName.$lastName@$emailDomain","SMTP:$username@$emailDomain")
    }

    # Fjern spesialtegn fra proxy-adresser
    $proxyAddresses = $proxyAddresses | ForEach-Object {
        $_.Replace("ø", "o").Replace("æ", "ae").Replace("å", "a").Replace(" ", "").ToLower()
    }

    # Generer et passord
    $uppercase = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    $lowercase = 'abcdefghijklmnopqrstuvwxyz'
    $password = (Get-Random -InputObject $uppercase[(Get-Random -Minimum 0 -Maximum 25)]).ToString() + `
                (Get-Random -InputObject $lowercase[(Get-Random -Minimum 0 -Maximum 25)]).ToString() + `
                (Get-Random -InputObject $lowercase[(Get-Random -Minimum 0 -Maximum 25)]).ToString() + `
                (Get-Random -Minimum 10000 -Maximum 99999).ToString()

    try {
        # Opprett ny bruker basert på mal
        New-ADUser `
            -Name "$name" `
            -SamAccountName $username `
            -UserPrincipalName $email `
            -EmailAddress $email `
            -Enabled $true `
            -GivenName $firstName `
            -Surname $lastName `
            -Instance $template `
            -Path $ou `
            -AccountPassword (ConvertTo-SecureString $password -AsPlainText -Force) `
            -OtherAttributes @{'proxyAddresses'=$proxyAddresses}

        # Kopier gruppemedlemskap fra mal
        $CopyFromUser = Get-ADUser $template -Properties MemberOf
        $CopyToUser = Get-ADUser $username -Properties MemberOf
        $CopyFromUser.MemberOf | Where-Object {$CopyToUser.MemberOf -notcontains $_} | ForEach-Object {
            Add-ADGroupMember -Identity $_ -Members $CopyToUser
        }

        # Legg til bruker i spesifikk gruppe
        Add-ADPrincipalGroupMembership -Identity $username -MemberOf 'G_WVD-Indresone-Brukere'

        # Vis suksessmelding
        [System.Windows.Forms.Clipboard]::SetText($password)
        [System.Windows.Forms.MessageBox]::Show("Brukeren $name er opprettet med brukernavn: $username og passord: $password under $department. Passordet er kopiert til utklippstavlen.")
        $name, $email, $password | Out-Host

    } catch {
        # Vis feilmelding
        [System.Windows.Forms.MessageBox]::Show("Feil ved oppretting av ny bruker: $($_.Exception.Message)")
    }

    # Lukk formen
    $form.Close()
})
$form.Controls.Add($createButton)

# Vis formen
$form.ShowDialog()
