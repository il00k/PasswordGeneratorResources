# Add required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Global Variables
$PasswordHistory = @()
$ErrorLogPath = "$env:TEMP\PasswordGeneratorErrors.log"
$ScriptVersion = '1.0.1' # Update this version number as needed
$CustomWords = @()
$LanguageResources = @{}
$SelectedLanguage = 'English'
$Themes = @{
    'Light' = @{
        'BackColor' = [System.Drawing.Color]::White
        'ForeColor' = [System.Drawing.Color]::Black
    }
    'Dark' = @{
        'BackColor' = [System.Drawing.Color]::Black
        'ForeColor' = [System.Drawing.Color]::White
    }
    'Blue' = @{
        'BackColor' = [System.Drawing.Color]::LightBlue
        'ForeColor' = [System.Drawing.Color]::Black
    }
}
$SelectedTheme = 'Light'

# Function to Log Errors
function Log-Error {
    param([string]$ErrorMessage)
    $Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $LogEntry = "[$Timestamp] ERROR: $ErrorMessage`n"
    Add-Content -Path $ErrorLogPath -Value $LogEntry
}

# Password Generation Functions

function Generate-Password {
    param(
        [int]$Length,
        [bool]$IncludeUppercase,
        [bool]$IncludeLowercase,
        [bool]$IncludeNumbers,
        [bool]$IncludeSpecialChars,
        [bool]$ExcludeSimilar,
        [string]$CustomSpecialChars
    )

    try {
        # Character Sets
        $CharSet = @()

        if ($IncludeUppercase) {
            $Uppercase = [char[]](65..90) # A-Z
            $CharSet += $Uppercase
        }
        if ($IncludeLowercase) {
            $Lowercase = [char[]](97..122) # a-z
            $CharSet += $Lowercase
        }
        if ($IncludeNumbers) {
            $Numbers = [char[]](48..57) # 0-9
            $CharSet += $Numbers
        }
        if ($IncludeSpecialChars) {
            if ([string]::IsNullOrEmpty($CustomSpecialChars)) {
                $SpecialChars = '!@#$%^&*()-_=+[]{}|;:",.<>/?`~'
            } else {
                $SpecialChars = $CustomSpecialChars
            }
            $CharSet += $SpecialChars.ToCharArray()
        }

        if ($ExcludeSimilar) {
            $SimilarChars = 'O','0','I','l','1'
            $CharSet = $CharSet | Where-Object { $SimilarChars -notcontains $_ }
        }

        if ($CharSet.Count -eq 0) {
            throw "No character sets selected."
        }

        $Password = -join ((1..$Length) | ForEach-Object { $CharSet | Get-Random })
        return $Password
    } catch {
        Log-Error $_.Exception.Message
        return "Error: $_"
    }
}

function Generate-Passphrase {
    param(
        [int]$WordCount,
        [string]$Delimiter,
        [array]$WordList
    )

    try {
        if ($WordList -eq $null -or $WordList.Count -eq 0) {
            # Load default word list from GitHub
            $WordListUrl = 'https://raw.githubusercontent.com/YourGitHubUsername/PasswordGeneratorResources/main/WordLists/custom_word_list.txt'
            try {
                $WordListContent = (Invoke-WebRequest -Uri $WordListUrl -UseBasicParsing).Content
                $WordList = $WordListContent -split "`r`n"
            } catch {
                Log-Error $_.Exception.Message
                # Use a simple built-in word list as a fallback
                $WordList = @(
                    'apple','banana','cherry','date','elderberry','fig','grape','honeydew',
                    'kiwi','lemon','mango','nectarine','orange','papaya','quince','raspberry',
                    'strawberry','tangerine','ugli','violet','watermelon','xigua','yam','zucchini'
                )
            }
        }
        $Passphrase = (1..$WordCount | ForEach-Object { $WordList | Get-Random }) -join $Delimiter
        return $Passphrase
    } catch {
        Log-Error $_.Exception.Message
        return "Error: $_"
    }
}

function Generate-PronounceablePassword {
    param([int]$Length)
    $Consonants = 'bcdfghjklmnpqrstvwxyz'
    $Vowels = 'aeiou'
    $Password = ''
    for ($i = 0; $i -lt $Length; $i++) {
        if ($i % 2 -eq 0) {
            $Password += $Consonants[(Get-Random -Maximum $Consonants.Length)]
        } else {
            $Password += $Vowels[(Get-Random -Maximum $Vowels.Length)]
        }
    }
    return $Password
}

function Get-PasswordStrength {
    param([string]$Password)

    try {
        $LengthScore = Switch ($Password.Length) {
            {$_ -ge 16} {3}
            {$_ -ge 12} {2}
            {$_ -ge 8}  {1}
            Default     {0}
        }

        $ComplexityScore = 0
        if ($Password -match '[a-z]') { $ComplexityScore++ }
        if ($Password -match '[A-Z]') { $ComplexityScore++ }
        if ($Password -match '\d')    { $ComplexityScore++ }
        if ($Password -match '[^a-zA-Z0-9]') { $ComplexityScore++ }

        $TotalScore = $LengthScore + $ComplexityScore

        Switch ($TotalScore) {
            {$_ -le 3}  { return 'Weak' }
            {$_ -le 5}  { return 'Moderate' }
            {$_ -le 7}  { return 'Strong' }
            {$_ -ge 8}  { return 'Very Strong' }
        }
    } catch {
        Log-Error $_.Exception.Message
        return "Unknown"
    }
}

# Check for Updates using GitHub
function Check-ForUpdates {
    try {
        $LatestVersionUrl = 'https://raw.githubusercontent.com/YourGitHubUsername/PasswordGeneratorResources/main/latest_version.txt'
        $LatestVersion = (Invoke-WebRequest -Uri $LatestVersionUrl -UseBasicParsing).Content.Trim()
        if ($LatestVersion -ne $ScriptVersion) {
            $UpdateMessage = "A new version ($LatestVersion) is available."
            [System.Windows.Forms.MessageBox]::Show($UpdateMessage, 'Update Available', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    } catch {
        Log-Error $_.Exception.Message
    }
}

# Load Language Resources
function Load-LanguageResources {
    param([string]$Language)
    $ResourceUrl = "https://raw.githubusercontent.com/YourGitHubUsername/PasswordGeneratorResources/main/Languages/$Language.json"
    try {
        $LanguageJson = (Invoke-WebRequest -Uri $ResourceUrl -UseBasicParsing).Content
        $LanguageResources = $LanguageJson | ConvertFrom-Json
    } catch {
        Log-Error $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show("Language file not found.", 'Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# Call the update checker and load default language on startup
Check-ForUpdates
Load-LanguageResources -Language $SelectedLanguage

# Create the form
$Form = New-Object System.Windows.Forms.Form
$Form.Text = 'Password Generator'
$Form.Size = New-Object System.Drawing.Size(500, 900)
$Form.StartPosition = "CenterScreen"
$Form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$Form.BackColor = $Themes[$SelectedTheme]['BackColor']
$Form.ForeColor = $Themes[$SelectedTheme]['ForeColor']
$Form.MaximizeBox = $true
$Form.FormBorderStyle = 'Sizable'

# Controls
# Password Length Label and Numeric UpDown
$LengthLabel = New-Object System.Windows.Forms.Label
$LengthLabel.Text = $LanguageResources.PasswordLength
$LengthLabel.Location = New-Object System.Drawing.Point(10, 20)
$LengthLabel.AutoSize = $true

$LengthUpDown = New-Object System.Windows.Forms.NumericUpDown
$LengthUpDown.Location = New-Object System.Drawing.Point(150, 18)
$LengthUpDown.Minimum = 1
$LengthUpDown.Maximum = 128
$LengthUpDown.Value = 12

# Include Uppercase Checkbox
$UppercaseCheckBox = New-Object System.Windows.Forms.CheckBox
$UppercaseCheckBox.Text = $LanguageResources.IncludeUppercase
$UppercaseCheckBox.Location = New-Object System.Drawing.Point(10, 60)
$UppercaseCheckBox.Checked = $true
$UppercaseCheckBox.AutoSize = $true

# Include Lowercase Checkbox
$LowercaseCheckBox = New-Object System.Windows.Forms.CheckBox
$LowercaseCheckBox.Text = $LanguageResources.IncludeLowercase
$LowercaseCheckBox.Location = New-Object System.Drawing.Point(10, 85)
$LowercaseCheckBox.Checked = $true
$LowercaseCheckBox.AutoSize = $true

# Include Numbers Checkbox
$NumbersCheckBox = New-Object System.Windows.Forms.CheckBox
$NumbersCheckBox.Text = $LanguageResources.IncludeNumbers
$NumbersCheckBox.Location = New-Object System.Drawing.Point(10, 110)
$NumbersCheckBox.Checked = $true
$NumbersCheckBox.AutoSize = $true

# Include Special Characters Checkbox
$SpecialCharsCheckBox = New-Object System.Windows.Forms.CheckBox
$SpecialCharsCheckBox.Text = $LanguageResources.IncludeSpecialChars
$SpecialCharsCheckBox.Location = New-Object System.Drawing.Point(10, 135)
$SpecialCharsCheckBox.Checked = $true
$SpecialCharsCheckBox.AutoSize = $true

# Exclude Similar Characters Checkbox
$ExcludeSimilarCheckBox = New-Object System.Windows.Forms.CheckBox
$ExcludeSimilarCheckBox.Text = $LanguageResources.ExcludeSimilarChars
$ExcludeSimilarCheckBox.Location = New-Object System.Drawing.Point(10, 160)
$ExcludeSimilarCheckBox.Checked = $false
$ExcludeSimilarCheckBox.AutoSize = $true

# Custom Special Characters Label and TextBox
$CustomSpecialCharsLabel = New-Object System.Windows.Forms.Label
$CustomSpecialCharsLabel.Text = $LanguageResources.CustomSpecialChars
$CustomSpecialCharsLabel.Location = New-Object System.Drawing.Point(10, 190)
$CustomSpecialCharsLabel.AutoSize = $true

$CustomSpecialCharsTextBox = New-Object System.Windows.Forms.TextBox
$CustomSpecialCharsTextBox.Location = New-Object System.Drawing.Point(150, 188)
$CustomSpecialCharsTextBox.Size = New-Object System.Drawing.Size(270, 20)
$CustomSpecialCharsTextBox.Text = ''

# Mask Password Checkbox
$MaskPasswordCheckBox = New-Object System.Windows.Forms.CheckBox
$MaskPasswordCheckBox.Text = $LanguageResources.MaskPassword
$MaskPasswordCheckBox.Location = New-Object System.Drawing.Point(10, 220)
$MaskPasswordCheckBox.Checked = $false
$MaskPasswordCheckBox.AutoSize = $true

# Generate Multiple Passwords Checkbox
$MultiplePasswordsCheckBox = New-Object System.Windows.Forms.CheckBox
$MultiplePasswordsCheckBox.Text = 'Generate Multiple Passwords'
$MultiplePasswordsCheckBox.Location = New-Object System.Drawing.Point(10, 245)
$MultiplePasswordsCheckBox.Checked = $false
$MultiplePasswordsCheckBox.AutoSize = $true

# Number of Passwords Label and Numeric UpDown
$NumberOfPasswordsLabel = New-Object System.Windows.Forms.Label
$NumberOfPasswordsLabel.Text = 'Number of Passwords:'
$NumberOfPasswordsLabel.Location = New-Object System.Drawing.Point(10, 270)
$NumberOfPasswordsLabel.AutoSize = $true
$NumberOfPasswordsLabel.Enabled = $false

$NumberOfPasswordsUpDown = New-Object System.Windows.Forms.NumericUpDown
$NumberOfPasswordsUpDown.Location = New-Object System.Drawing.Point(150, 268)
$NumberOfPasswordsUpDown.Minimum = 1
$NumberOfPasswordsUpDown.Maximum = 100
$NumberOfPasswordsUpDown.Value = 5
$NumberOfPasswordsUpDown.Enabled = $false

# Passphrase Generation Checkbox
$PassphraseCheckBox = New-Object System.Windows.Forms.CheckBox
$PassphraseCheckBox.Text = $LanguageResources.GeneratePassphrase
$PassphraseCheckBox.Location = New-Object System.Drawing.Point(10, 295)
$PassphraseCheckBox.Checked = $false
$PassphraseCheckBox.AutoSize = $true

# Number of Words Label and Numeric UpDown
$NumberOfWordsLabel = New-Object System.Windows.Forms.Label
$NumberOfWordsLabel.Text = $LanguageResources.NumberOfWords
$NumberOfWordsLabel.Location = New-Object System.Drawing.Point(10, 320)
$NumberOfWordsLabel.AutoSize = $true
$NumberOfWordsLabel.Enabled = $false

$NumberOfWordsUpDown = New-Object System.Windows.Forms.NumericUpDown
$NumberOfWordsUpDown.Location = New-Object System.Drawing.Point(150, 318)
$NumberOfWordsUpDown.Minimum = 2
$NumberOfWordsUpDown.Maximum = 10
$NumberOfWordsUpDown.Value = 4
$NumberOfWordsUpDown.Enabled = $false

# Delimiter Label and TextBox
$DelimiterLabel = New-Object System.Windows.Forms.Label
$DelimiterLabel.Text = $LanguageResources.Delimiter
$DelimiterLabel.Location = New-Object System.Drawing.Point(10, 345)
$DelimiterLabel.AutoSize = $true
$DelimiterLabel.Enabled = $false

$DelimiterTextBox = New-Object System.Windows.Forms.TextBox
$DelimiterTextBox.Location = New-Object System.Drawing.Point(150, 343)
$DelimiterTextBox.Size = New-Object System.Drawing.Size(50, 20)
$DelimiterTextBox.Text = '-'
$DelimiterTextBox.Enabled = $false

# Load Custom Word List Button
$LoadWordListButton = New-Object System.Windows.Forms.Button
$LoadWordListButton.Text = $LanguageResources.LoadCustomWordList
$LoadWordListButton.Location = New-Object System.Drawing.Point(220, 318)
$LoadWordListButton.Size = New-Object System.Drawing.Size(200, 30)
$LoadWordListButton.Enabled = $false

# Pronounceable Password Checkbox
$PronounceableCheckBox = New-Object System.Windows.Forms.CheckBox
$PronounceableCheckBox.Text = $LanguageResources.PronounceablePassword
$PronounceableCheckBox.Location = New-Object System.Drawing.Point(10, 375)
$PronounceableCheckBox.Checked = $false
$PronounceableCheckBox.AutoSize = $true

# Generate Password Button
$GenerateButton = New-Object System.Windows.Forms.Button
$GenerateButton.Text = $LanguageResources.GeneratePassword
$GenerateButton.Location = New-Object System.Drawing.Point(10, 410)
$GenerateButton.Size = New-Object System.Drawing.Size(200, 30)

# Copy to Clipboard Button
$CopyButton = New-Object System.Windows.Forms.Button
$CopyButton.Text = $LanguageResources.CopyToClipboard
$CopyButton.Location = New-Object System.Drawing.Point(220, 410)
$CopyButton.Size = New-Object System.Drawing.Size(200, 30)

# Password Strength Label
$StrengthLabel = New-Object System.Windows.Forms.Label
$StrengthLabel.Text = $LanguageResources.PasswordStrength
$StrengthLabel.Location = New-Object System.Drawing.Point(10, 450)
$StrengthLabel.AutoSize = $true

# Password Strength Progress Bar
$StrengthProgressBar = New-Object System.Windows.Forms.ProgressBar
$StrengthProgressBar.Location = New-Object System.Drawing.Point(150, 450)
$StrengthProgressBar.Size = New-Object System.Drawing.Size(270, 20)
$StrengthProgressBar.Minimum = 0
$StrengthProgressBar.Maximum = 100

# Generated Password Label and TextBox
$PasswordLabel = New-Object System.Windows.Forms.Label
$PasswordLabel.Text = $LanguageResources.GeneratedPassword
$PasswordLabel.Location = New-Object System.Drawing.Point(10, 480)
$PasswordLabel.AutoSize = $true

$PasswordTextBox = New-Object System.Windows.Forms.TextBox
$PasswordTextBox.Location = New-Object System.Drawing.Point(10, 500)
$PasswordTextBox.Size = New-Object System.Drawing.Size(460, 80)
$PasswordTextBox.ReadOnly = $true
$PasswordTextBox.Multiline = $true
$PasswordTextBox.ScrollBars = 'Vertical'
$PasswordTextBox.UseSystemPasswordChar = $MaskPasswordCheckBox.Checked

# Password History Button
$HistoryButton = New-Object System.Windows.Forms.Button
$HistoryButton.Text = $LanguageResources.PasswordHistory
$HistoryButton.Location = New-Object System.Drawing.Point(10, 590)
$HistoryButton.Size = New-Object System.Drawing.Size(130, 30)

# Save Password Button
$SaveButton = New-Object System.Windows.Forms.Button
$SaveButton.Text = $LanguageResources.SavePassword
$SaveButton.Location = New-Object System.Drawing.Point(150, 590)
$SaveButton.Size = New-Object System.Drawing.Size(130, 30)

# Set Expiration Reminder Button
$SetReminderButton = New-Object System.Windows.Forms.Button
$SetReminderButton.Text = $LanguageResources.SetExpirationReminder
$SetReminderButton.Location = New-Object System.Drawing.Point(290, 590)
$SetReminderButton.Size = New-Object System.Drawing.Size(180, 30)

# Theme Selection Label and Dropdown
$ThemeLabel = New-Object System.Windows.Forms.Label
$ThemeLabel.Text = $LanguageResources.SelectTheme
$ThemeLabel.Location = New-Object System.Drawing.Point(10, 630)
$ThemeLabel.AutoSize = $true

$ThemeDropdown = New-Object System.Windows.Forms.ComboBox
$ThemeDropdown.Location = New-Object System.Drawing.Point(150, 628)
$ThemeDropdown.Size = New-Object System.Drawing.Size(270, 20)
$ThemeDropdown.Items.AddRange($Themes.Keys)
$ThemeDropdown.SelectedItem = $SelectedTheme

# Language Selection Label and Dropdown
$LanguageLabel = New-Object System.Windows.Forms.Label
$LanguageLabel.Text = $LanguageResources.Language
$LanguageLabel.Location = New-Object System.Drawing.Point(10, 660)
$LanguageLabel.AutoSize = $true

$LanguageDropdown = New-Object System.Windows.Forms.ComboBox
$LanguageDropdown.Location = New-Object System.Drawing.Point(150, 658)
$LanguageDropdown.Size = New-Object System.Drawing.Size(270, 20)
$LanguageDropdown.Items.AddRange(@('English', 'Spanish', 'French'))
$LanguageDropdown.SelectedItem = $SelectedLanguage

# Export to CSV Button
$ExportCSVButton = New-Object System.Windows.Forms.Button
$ExportCSVButton.Text = 'Export to CSV'
$ExportCSVButton.Location = New-Object System.Drawing.Point(10, 700)
$ExportCSVButton.Size = New-Object System.Drawing.Size(130, 30)

# Export to JSON Button
$ExportJSONButton = New-Object System.Windows.Forms.Button
$ExportJSONButton.Text = 'Export to JSON'
$ExportJSONButton.Location = New-Object System.Drawing.Point(150, 700)
$ExportJSONButton.Size = New-Object System.Drawing.Size(130, 30)

# Add controls to the form
$Form.Controls.Add($LengthLabel)
$Form.Controls.Add($LengthUpDown)
$Form.Controls.Add($UppercaseCheckBox)
$Form.Controls.Add($LowercaseCheckBox)
$Form.Controls.Add($NumbersCheckBox)
$Form.Controls.Add($SpecialCharsCheckBox)
$Form.Controls.Add($ExcludeSimilarCheckBox)
$Form.Controls.Add($CustomSpecialCharsLabel)
$Form.Controls.Add($CustomSpecialCharsTextBox)
$Form.Controls.Add($MaskPasswordCheckBox)
$Form.Controls.Add($MultiplePasswordsCheckBox)
$Form.Controls.Add($NumberOfPasswordsLabel)
$Form.Controls.Add($NumberOfPasswordsUpDown)
$Form.Controls.Add($PassphraseCheckBox)
$Form.Controls.Add($NumberOfWordsLabel)
$Form.Controls.Add($NumberOfWordsUpDown)
$Form.Controls.Add($DelimiterLabel)
$Form.Controls.Add($DelimiterTextBox)
$Form.Controls.Add($LoadWordListButton)
$Form.Controls.Add($PronounceableCheckBox)
$Form.Controls.Add($GenerateButton)
$Form.Controls.Add($CopyButton)
$Form.Controls.Add($StrengthLabel)
$Form.Controls.Add($StrengthProgressBar)
$Form.Controls.Add($PasswordLabel)
$Form.Controls.Add($PasswordTextBox)
$Form.Controls.Add($HistoryButton)
$Form.Controls.Add($SaveButton)
$Form.Controls.Add($SetReminderButton)
$Form.Controls.Add($ThemeLabel)
$Form.Controls.Add($ThemeDropdown)
$Form.Controls.Add($LanguageLabel)
$Form.Controls.Add($LanguageDropdown)
$Form.Controls.Add($ExportCSVButton)
$Form.Controls.Add($ExportJSONButton)

# Event Handlers
$MultiplePasswordsCheckBox.Add_CheckedChanged({
    $NumberOfPasswordsLabel.Enabled = $MultiplePasswordsCheckBox.Checked
    $NumberOfPasswordsUpDown.Enabled = $MultiplePasswordsCheckBox.Checked
})

$PassphraseCheckBox.Add_CheckedChanged({
    $IsPassphrase = $PassphraseCheckBox.Checked
    $LengthLabel.Enabled = -not $IsPassphrase
    $LengthUpDown.Enabled = -not $IsPassphrase
    $UppercaseCheckBox.Enabled = -not $IsPassphrase
    $LowercaseCheckBox.Enabled = -not $IsPassphrase
    $NumbersCheckBox.Enabled = -not $IsPassphrase
    $SpecialCharsCheckBox.Enabled = -not $IsPassphrase
    $ExcludeSimilarCheckBox.Enabled = -not $IsPassphrase
    $CustomSpecialCharsLabel.Enabled = -not $IsPassphrase
    $CustomSpecialCharsTextBox.Enabled = -not $IsPassphrase
    $NumberOfWordsLabel.Enabled = $IsPassphrase
    $NumberOfWordsUpDown.Enabled = $IsPassphrase
    $DelimiterLabel.Enabled = $IsPassphrase
    $DelimiterTextBox.Enabled = $IsPassphrase
    $LoadWordListButton.Enabled = $IsPassphrase
})

$LoadWordListButton.Add_Click({
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Filter = "Text Files (*.txt)|*.txt"
    if ($OpenFileDialog.ShowDialog() -eq 'OK') {
        try {
            $CustomWords = Get-Content $OpenFileDialog.FileName
        } catch {
            Log-Error $_.Exception.Message
            [System.Windows.Forms.MessageBox]::Show($LanguageResources.FailedToLoadWordList, $LanguageResources.Error, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

$PronounceableCheckBox.Add_CheckedChanged({
    $IsPronounceable = $PronounceableCheckBox.Checked
    # Disable other password options when pronounceable password is selected
    $LengthLabel.Enabled = -not $IsPronounceable
    $LengthUpDown.Enabled = -not $IsPronounceable
    $UppercaseCheckBox.Enabled = -not $IsPronounceable
    $LowercaseCheckBox.Enabled = -not $IsPronounceable
    $NumbersCheckBox.Enabled = -not $IsPronounceable
    $SpecialCharsCheckBox.Enabled = -not $IsPronounceable
    $ExcludeSimilarCheckBox.Enabled = -not $IsPronounceable
    $CustomSpecialCharsLabel.Enabled = -not $IsPronounceable
    $CustomSpecialCharsTextBox.Enabled = -not $IsPronounceable
    $PassphraseCheckBox.Enabled = -not $IsPronounceable
})

$GenerateButton.Add_Click({
    try {
        $Passwords = @()
        $Strengths = @()

        $Count = 1
        if ($MultiplePasswordsCheckBox.Checked) {
            $Count = [int]$NumberOfPasswordsUpDown.Value
        }

        for ($i = 0; $i -lt $Count; $i++) {
            if ($PassphraseCheckBox.Checked) {
                $Password = Generate-Passphrase -WordCount $NumberOfWordsUpDown.Value -Delimiter $DelimiterTextBox.Text -WordList $CustomWords
                $Strengths += 'Passphrase'
            } elseif ($PronounceableCheckBox.Checked) {
                $Password = Generate-PronounceablePassword -Length $LengthUpDown.Value
                $Strengths += Get-PasswordStrength -Password $Password
            } else {
                $Password = Generate-Password -Length $LengthUpDown.Value `
                    -IncludeUppercase $UppercaseCheckBox.Checked `
                    -IncludeLowercase $LowercaseCheckBox.Checked `
                    -IncludeNumbers $NumbersCheckBox.Checked `
                    -IncludeSpecialChars $SpecialCharsCheckBox.Checked `
                    -ExcludeSimilar $ExcludeSimilarCheckBox.Checked `
                    -CustomSpecialChars $CustomSpecialCharsTextBox.Text
                $Strengths += Get-PasswordStrength -Password $Password
            }

            if ($Password -like "Error*") {
                [System.Windows.Forms.MessageBox]::Show($Password, $LanguageResources.Error, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            } else {
                $Passwords += $Password
                $PasswordHistory += $Password
            }
        }

        $PasswordTextBox.Text = $Passwords -join "`r`n"
        $UniqueStrengths = $Strengths | Select-Object -Unique
        $StrengthLabel.Text = "$($LanguageResources.PasswordStrength) " + ($UniqueStrengths -join ', ')

        # Update Password Strength Progress Bar
        switch ($UniqueStrengths[0]) {
            'Weak' {
                $StrengthProgressBar.Value = 25
                $StrengthProgressBar.ForeColor = 'Red'
            }
            'Moderate' {
                $StrengthProgressBar.Value = 50
                $StrengthProgressBar.ForeColor = 'Orange'
            }
            'Strong' {
                $StrengthProgressBar.Value = 75
                $StrengthProgressBar.ForeColor = 'YellowGreen'
            }
            'Very Strong' {
                $StrengthProgressBar.Value = 100
                $StrengthProgressBar.ForeColor = 'Green'
            }
            Default {
                $StrengthProgressBar.Value = 0
            }
        }
    } catch {
        Log-Error $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show($LanguageResources.FailedToGenerate, $LanguageResources.Error, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$CopyButton.Add_Click({
    try {
        [System.Windows.Forms.Clipboard]::SetText($PasswordTextBox.Text)
        [System.Windows.Forms.MessageBox]::Show($LanguageResources.PasswordCopied, $LanguageResources.Info, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

        # Clipboard Clear Timer
        if ($null -ne $Timer) {
            $Timer.Stop()
            $Timer.Dispose()
        }

        $Timer = New-Object System.Windows.Forms.Timer
        $Timer.Interval = 10000 # 10 seconds
        $Timer.Add_Tick({
            try {
                [System.Windows.Forms.Clipboard]::Clear()
                $Timer.Stop()
                $Timer.Dispose()
                [System.Windows.Forms.MessageBox]::Show($LanguageResources.ClipboardCleared, $LanguageResources.Info, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } catch {
                Log-Error $_.Exception.Message
            }
        })
        $Timer.Start()
    } catch {
        Log-Error $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show($LanguageResources.FailedToCopy, $LanguageResources.Error, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$MaskPasswordCheckBox.Add_CheckedChanged({
    $PasswordTextBox.UseSystemPasswordChar = $MaskPasswordCheckBox.Checked
})

$HistoryButton.Add_Click({
    try {
        $HistoryForm = New-Object System.Windows.Forms.Form
        $HistoryForm.Text = $LanguageResources.PasswordHistory
        $HistoryForm.Size = New-Object System.Drawing.Size(400, 300)
        $HistoryForm.StartPosition = "CenterParent"
        $HistoryForm.FormBorderStyle = 'Sizable'
        $HistoryForm.MaximizeBox = $true

        $HistoryTextBox = New-Object System.Windows.Forms.TextBox
        $HistoryTextBox.Location = New-Object System.Drawing.Point(10, 10)
        $HistoryTextBox.Size = New-Object System.Drawing.Size(360, 240)
        $HistoryTextBox.Multiline = $true
        $HistoryTextBox.ScrollBars = 'Vertical'
        $HistoryTextBox.ReadOnly = $true
        $HistoryTextBox.Text = ($PasswordHistory -join "`r`n")
        $HistoryTextBox.UseSystemPasswordChar = $MaskPasswordCheckBox.Checked
        $HistoryTextBox.Anchor = 'Top, Left, Right, Bottom'

        $HistoryForm.Controls.Add($HistoryTextBox)
        [void]$HistoryForm.ShowDialog()
    } catch {
        Log-Error $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show($LanguageResources.FailedToDisplayHistory, $LanguageResources.Error, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$SaveButton.Add_Click({
    try {
        if ([string]::IsNullOrEmpty($PasswordTextBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show($LanguageResources.NoPasswordToSave, $LanguageResources.Info, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }

        # Prompt for passphrase
        $PassphraseForm = New-Object System.Windows.Forms.Form
        $PassphraseForm.Text = 'Enter Passphrase'
        $PassphraseForm.Size = New-Object System.Drawing.Size(300, 150)
        $PassphraseForm.StartPosition = "CenterParent"

        $PassphraseLabel = New-Object System.Windows.Forms.Label
        $PassphraseLabel.Text = 'Passphrase:'
        $PassphraseLabel.Location = New-Object System.Drawing.Point(10, 20)
        $PassphraseLabel.AutoSize = $true

        $PassphraseTextBox = New-Object System.Windows.Forms.TextBox
        $PassphraseTextBox.Location = New-Object System.Drawing.Point(100, 18)
        $PassphraseTextBox.Size = New-Object System.Drawing.Size(170, 20)
        $PassphraseTextBox.UseSystemPasswordChar = $true

        $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Text = 'OK'
        $OKButton.Location = New-Object System.Drawing.Point(100, 50)
        $OKButton.Size = New-Object System.Drawing.Size(80, 30)

        $PassphraseForm.Controls.Add($PassphraseLabel)
        $PassphraseForm.Controls.Add($PassphraseTextBox)
        $PassphraseForm.Controls.Add($OKButton)

        $OKButton.Add_Click({
            $Passphrase = $PassphraseTextBox.Text
            $PassphraseForm.Close()
        })

        [void]$PassphraseForm.ShowDialog()

        if ([string]::IsNullOrEmpty($Passphrase)) {
            [System.Windows.Forms.MessageBox]::Show('Passphrase is required.', $LanguageResources.Error, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        # Encrypt the password
        $SecureString = ConvertTo-SecureString -String $PasswordTextBox.Text -AsPlainText -Force
        $EncryptedPassword = $SecureString | ConvertFrom-SecureString -SecureKey (ConvertTo-SecureString -String $Passphrase -AsPlainText -Force)

        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.Filter = "Encrypted Files (*.enc)|*.enc"
        if ($SaveFileDialog.ShowDialog() -eq 'OK') {
            $EncryptedPassword | Out-File -FilePath $SaveFileDialog.FileName -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show($LanguageResources.PasswordSaved, $LanguageResources.Info, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    } catch {
        Log-Error $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show($LanguageResources.FailedToSave, $LanguageResources.Error, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$SetReminderButton.Add_Click({
    try {
        $ReminderForm = New-Object System.Windows.Forms.Form
        $ReminderForm.Text = $LanguageResources.SetExpirationReminder
        $ReminderForm.Size = New-Object System.Drawing.Size(300, 150)
        $ReminderForm.StartPosition = "CenterParent"

        $DateLabel = New-Object System.Windows.Forms.Label
        $DateLabel.Text = 'Select Expiration Date:'
        $DateLabel.Location = New-Object System.Drawing.Point(10, 20)
        $DateLabel.AutoSize = $true

        $DateTimePicker = New-Object System.Windows.Forms.DateTimePicker
        $DateTimePicker.Location = New-Object System.Drawing.Point(10, 50)
        $DateTimePicker.Format = 'Short'

        $SetButton = New-Object System.Windows.Forms.Button
        $SetButton.Text = 'Set Reminder'
        $SetButton.Location = New-Object System.Drawing.Point(10, 80)
        $SetButton.Size = New-Object System.Drawing.Size(100, 30)

        $ReminderForm.Controls.Add($DateLabel)
        $ReminderForm.Controls.Add($DateTimePicker)
        $ReminderForm.Controls.Add($SetButton)

        $SetButton.Add_Click({
            try {
                $TriggerDate = $DateTimePicker.Value
                if ($TriggerDate -le (Get-Date)) {
                    [System.Windows.Forms.MessageBox]::Show('Please select a future date.', $LanguageResources.Error, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    return
                }

                $Trigger = New-JobTrigger -Once -At $TriggerDate
                $Options = New-ScheduledJobOption -RunElevated

                Register-ScheduledJob -Name "PasswordExpirationReminder" -Trigger $Trigger -ScheduledJobOption $Options -ScriptBlock {
                    [System.Windows.Forms.MessageBox]::Show("It's time to change your password.", 'Password Expiration Reminder', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                }

                [System.Windows.Forms.MessageBox]::Show('Password expiration reminder set.', $LanguageResources.Info, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                $ReminderForm.Close()
            } catch {
                Log-Error $_.Exception.Message
                [System.Windows.Forms.MessageBox]::Show("Failed to set reminder.", $LanguageResources.Error, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        })

        [void]$ReminderForm.ShowDialog()
    } catch {
        Log-Error $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show("Failed to open reminder settings.", $LanguageResources.Error, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$ThemeDropdown.Add_SelectedIndexChanged({
    $SelectedTheme = $ThemeDropdown.SelectedItem
    $Form.BackColor = $Themes[$SelectedTheme]['BackColor']
    $Form.ForeColor = $Themes[$SelectedTheme]['ForeColor']
    # Update other controls' colors if necessary
})

$LanguageDropdown.Add_SelectedIndexChanged({
    $SelectedLanguage = $LanguageDropdown.SelectedItem
    Load-LanguageResources -Language $SelectedLanguage
    # Update all control texts with new translations
    $Form.Text = $LanguageResources.FormTitle
    $LengthLabel.Text = $LanguageResources.PasswordLength
    $UppercaseCheckBox.Text = $LanguageResources.IncludeUppercase
    $LowercaseCheckBox.Text = $LanguageResources.IncludeLowercase
    $NumbersCheckBox.Text = $LanguageResources.IncludeNumbers
    $SpecialCharsCheckBox.Text = $LanguageResources.IncludeSpecialChars
    $ExcludeSimilarCheckBox.Text = $LanguageResources.ExcludeSimilarChars
    $CustomSpecialCharsLabel.Text = $LanguageResources.CustomSpecialChars
    $MaskPasswordCheckBox.Text = $LanguageResources.MaskPassword
    $GenerateButton.Text = $LanguageResources.GeneratePassword
    $CopyButton.Text = $LanguageResources.CopyToClipboard
    $StrengthLabel.Text = $LanguageResources.PasswordStrength
    $PasswordLabel.Text = $LanguageResources.GeneratedPassword
    $HistoryButton.Text = $LanguageResources.PasswordHistory
    $SaveButton.Text = $LanguageResources.SavePassword
    $SetReminderButton.Text = $LanguageResources.SetExpirationReminder
    $LoadWordListButton.Text = $LanguageResources.LoadCustomWordList
    $LanguageLabel.Text = $LanguageResources.Language
    $ThemeLabel.Text = $LanguageResources.SelectTheme
    $PassphraseCheckBox.Text = $LanguageResources.GeneratePassphrase
    $NumberOfWordsLabel.Text = $LanguageResources.NumberOfWords
    $DelimiterLabel.Text = $LanguageResources.Delimiter
    $PronounceableCheckBox.Text = $LanguageResources.PronounceablePassword
})

$ExportCSVButton.Add_Click({
    try {
        if ($PasswordHistory.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('No passwords to export.', $LanguageResources.Info, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
        if ($SaveFileDialog.ShowDialog() -eq 'OK') {
            $PasswordHistory | Export-Csv -Path $SaveFileDialog.FileName -NoTypeInformation
            [System.Windows.Forms.MessageBox]::Show('Passwords exported to CSV.', $LanguageResources.Info, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    } catch {
        Log-Error $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show("Failed to export passwords.", $LanguageResources.Error, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$ExportJSONButton.Add_Click({
    try {
        if ($PasswordHistory.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('No passwords to export.', $LanguageResources.Info, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.Filter = "JSON Files (*.json)|*.json"
        if ($SaveFileDialog.ShowDialog() -eq 'OK') {
            $PasswordHistory | ConvertTo-Json | Out-File -FilePath $SaveFileDialog.FileName -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show('Passwords exported to JSON.', $LanguageResources.Info, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    } catch {
        Log-Error $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show("Failed to export passwords.", $LanguageResources.Error, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Adjust the form size and make controls anchor properly
$Form.AutoSize = $true
$Form.AutoSizeMode = 'GrowAndShrink'

# Anchor controls to resize with the form
$PasswordTextBox.Anchor = 'Top, Left, Right, Bottom'
$GenerateButton.Anchor = 'Bottom, Left'
$CopyButton.Anchor = 'Bottom, Right'
$HistoryButton.Anchor = 'Bottom, Left'
$SaveButton.Anchor = 'Bottom'
$SetReminderButton.Anchor = 'Bottom, Right'

# Show the form with error handling
try {
    [void]$Form.ShowDialog()
} catch {
    Log-Error $_.Exception.Message
    [System.Windows.Forms.MessageBox]::Show("An unexpected error occurred.", $LanguageResources.Error, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}
