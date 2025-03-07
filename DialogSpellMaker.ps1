# DialogSpellGenerator.ps1 - GUI version of the spell generator

# Add Windows Forms assembly for dialog boxes
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to get the script/executable path reliably in both script and compiled contexts
function Get-ScriptDirectory {
    if ($PSScriptRoot) {
        # Running as non-compiled script
        return $PSScriptRoot
    }
    elseif ($MyInvocation.MyCommand.Path) {
        # Running as non-compiled script, older PowerShell versions
        return Split-Path -Parent $MyInvocation.MyCommand.Path
    }
    elseif ($script:MyInvocation.MyCommand.Path) {
        # Running as non-compiled script, older PowerShell versions
        return Split-Path -Parent $script:MyInvocation.MyCommand.Path
    }
    elseif ([System.IO.Path]::GetDirectoryName($MyInvocation.PSScriptRoot)) {
        # For PS2EXE compiled scripts
        return [System.IO.Path]::GetDirectoryName($MyInvocation.PSScriptRoot)
    }
    elseif ([System.IO.Path]::GetDirectoryName([System.Reflection.Assembly]::GetExecutingAssembly().Location)) {
        # For PS2EXE compiled scripts, another method
        return [System.IO.Path]::GetDirectoryName([System.Reflection.Assembly]::GetExecutingAssembly().Location)
    }
    else {
        # Fallback to current directory if everything else fails
        return (Get-Location).Path
    }
}

# Function to format duration text for display
function Format-DurationText {
    param (
        [string]$duration
    )
    
    # Return "Instant" and "Permanent" as is
    if ($duration -eq "Instant" -or $duration -eq "Permanent") {
        return $duration
    }
    
    # Handle numbered durations with underscores
    if ($duration -match "^(\d+)_(minute|hour|day|week|month|year)$") {
        $number = $matches[1]
        $unit = $matches[2]
        
        # Convert to plural if needed
        if ($number -ne "1") {
            $unit = "${unit}s"
        }
        
        return "$number $unit"
    }
    
    # Return original if no pattern matches
    return $duration
}

# External JSON data file paths
$baseDir = Get-ScriptDirectory

$spellDataPath = Join-Path -Path $baseDir "spell_data.json"
$synonymsPath = Join-Path -Path $baseDir "synonyms.json"
$timingPatternsPath = Join-Path -Path $baseDir "timing_patterns.json"
$spellDescriptionsPath = Join-Path -Path $baseDir "spell_descriptions.json"

# Load JSON data from files
# Import ElementMapper module
$elementMapperPath = Join-Path -Path $baseDir "ElementMapper.ps1"
. $elementMapperPath

# Load JSON data from files
$spellDataJson = Get-Content -Path $spellDataPath -Raw
$synonymsJson = Get-Content -Path $synonymsPath -Raw

# Load timing patterns for duration recognition
$timingPatterns = Get-Content -Path $timingPatternsPath -Raw | ConvertFrom-Json
Write-Host "Loaded timing patterns from $timingPatternsPath"

# Parse JSON data
$spellData = ConvertFrom-Json $spellDataJson

# Load spell descriptions with error handling
try {
    if (Test-Path -Path $spellDescriptionsPath) {
        $spellDescriptionsJson = Get-Content -Path $spellDescriptionsPath -Raw
        $spellDescriptions = ConvertFrom-Json $spellDescriptionsJson
        Write-Host "Loaded spell descriptions from $spellDescriptionsPath"
        
        # Extract template elements for mapping
        $templateElements = @()
        
        # Handle the new nested structure
        if ($spellDescriptions.PSObject.Properties.Name -contains "spoken_spell_table" -and 
            $spellDescriptions.spoken_spell_table.PSObject.Properties.Name -contains "effect_prefix") {
            
            # Loop through each effect in the effect_prefix
            foreach ($effect in $spellDescriptions.spoken_spell_table.effect_prefix.PSObject.Properties.Name) {
                $effectObj = $spellDescriptions.spoken_spell_table.effect_prefix.$effect
                
                if ($effectObj.PSObject.Properties.Name -contains "element_prefix") {
                    # Loop through each element in the element_prefix for this effect
                    foreach ($element in $effectObj.element_prefix.PSObject.Properties.Name) {
                        if ($templateElements -notcontains $element) {
                            $templateElements += $element
                        }
                    }
                }
            }
        }
        
        Write-Host "Extracted $(($templateElements | Measure-Object).Count) template elements for mapping" -ForegroundColor Cyan
    } else {
        Write-Host "Warning: Spell descriptions file not found at $spellDescriptionsPath" -ForegroundColor Yellow
        $spellDescriptions = [PSCustomObject]@{}
        $templateElements = @()
    }
} catch {
    Write-Host "Error loading spell descriptions: $_" -ForegroundColor Red
    $spellDescriptions = [PSCustomObject]@{}
    $templateElements = @()
}

# Convert synonyms from space-separated strings in JSON to arrays of strings
$synonymsRaw = ConvertFrom-Json $synonymsJson
$synonyms = @{}
# Ensure timingPatterns is initialized
if (-not $timingPatterns) {
    Write-Host "Warning: Timing patterns file not found or empty, using defaults" -ForegroundColor Yellow
    # Fallback to default timing patterns if file not found
    $timingPatterns = @{
        "duration_patterns" = @{
            "Instant" = @("instantly", "immediate", "right away")
            "1_minute" = @("(?:for\s+)?(?:a|one)\s+minute", "(?:for\s+)?1\s+minute", "(?:for\s+)?(?:a|one)\s+min")
            "5_minute" = @("(?:for\s+)?(?:five|5)\s+minutes?", "(?:for\s+)?(?:five|5)\s+mins?")
            "10_minute" = @("(?:for\s+)?(?:ten|10)\s+minutes?", "(?:for\s+)?(?:ten|10)\s+mins?")
            "30_minute" = @("(?:for\s+)?(?:thirty|half an hour|30)\s+minutes?", "(?:for\s+)?(?:thirty|half an hour|30)\s+mins?")
            "1_hour" = @("(?:for\s+)?(?:an|one|1)\s+hour", "(?:for\s+)?(?:an|one|1)\s+hr")
            "5_hour" = @("(?:for\s+)?(?:five|5)\s+hours", "(?:for\s+)?(?:five|5)\s+hrs")
            "24_hour" = @("(?:for\s+)?(?:a|one)\s+day", "(?:for\s+)?(?:twenty[ -]four|24)\s+hours", "(?:for\s+)?(?:twenty[ -]four|24)\s+hrs")
        }
    }
}

# Process each category in the synonyms JSON immediately
foreach ($category in $synonymsRaw.PSObject.Properties.Name) {
    $synonyms[$category] = @{}
    
    # Process each key in the category
    foreach ($key in $synonymsRaw.$category.PSObject.Properties.Name) {
        # Split the space-separated string into an array of strings
        $synonyms[$category][$key] = $synonymsRaw.$category.$key.Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries)
    }
}


# Function to create an input dialog box for numeric input
function Get-NumberInputDialog {
    param(
        [string]$title,
        [string]$message,
        [int]$defaultValue = 1
    )
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $title
    $form.Size = New-Object System.Drawing.Size(350, 200)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 20)
    $label.Size = New-Object System.Drawing.Size(300, 40)
    $label.Text = $message
    $form.Controls.Add($label)
    
    $numericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $numericUpDown.Location = New-Object System.Drawing.Point(10, 70)
    $numericUpDown.Size = New-Object System.Drawing.Size(100, 25)
    $numericUpDown.Minimum = 1
    $numericUpDown.Maximum = 100
    $numericUpDown.Value = $defaultValue
    $form.Controls.Add($numericUpDown)
    
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(75, 120)
    $okButton.Size = New-Object System.Drawing.Size(75, 25)
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($okButton)
    $form.AcceptButton = $okButton
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(175, 120)
    $cancelButton.Size = New-Object System.Drawing.Size(75, 25)
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($cancelButton)
    $form.CancelButton = $cancelButton
    
    $form.TopMost = $true
    $result = $form.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return [int]$numericUpDown.Value
    } else {
        return $null
    }
}

# Function to create a yes/no dialog
function Get-YesNoDialog {
    param(
        [string]$title,
        [string]$message
    )
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        $message,
        $title,
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    
    return ($result -eq [System.Windows.Forms.DialogResult]::Yes)
}

# Function to create a selection dialog for choosing from a list
function Get-SelectionDialog {
    param(
        [string]$title,
        [string]$message,
        [array]$options
    )
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $title
    $form.Size = New-Object System.Drawing.Size(400, 350)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 20)
    $label.Size = New-Object System.Drawing.Size(380, 40)
    $label.Text = $message
    $form.Controls.Add($label)
    
    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Location = New-Object System.Drawing.Point(10, 70)
    $listBox.Size = New-Object System.Drawing.Size(360, 180)
    $listBox.SelectionMode = "One"
    $listBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    
    foreach ($option in $options) {
        [void]$listBox.Items.Add($option)
    }
    
    $listBox.SelectedIndex = 0
    $form.Controls.Add($listBox)
    
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(100, 270)
    $okButton.Size = New-Object System.Drawing.Size(75, 25)
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($okButton)
    $form.AcceptButton = $okButton
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(200, 270)
    $cancelButton.Size = New-Object System.Drawing.Size(75, 25)
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($cancelButton)
    $form.CancelButton = $cancelButton
    
    $form.TopMost = $true
    $result = $form.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $listBox.SelectedItem
    } else {
        return $null
    }
}

# Function to create a text input dialog
function Get-TextInputDialog {
    param(
        [string]$title,
        [string]$message,
        [string]$defaultValue = ""
    )
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $title
    $form.Size = New-Object System.Drawing.Size(500, 300)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 20)
    $label.Size = New-Object System.Drawing.Size(460, 40)
    $label.Text = $message
    $form.Controls.Add($label)
    
    $textBox = New-Object System.Windows.Forms.RichTextBox
    $textBox.Location = New-Object System.Drawing.Point(10, 70)
    $textBox.Size = New-Object System.Drawing.Size(460, 100)
    $textBox.Text = $defaultValue
    $textBox.Multiline = $true
    $textBox.ScrollBars = "Vertical"
    $form.Controls.Add($textBox)
    
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(125, 220)
    $okButton.Size = New-Object System.Drawing.Size(75, 25)
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($okButton)
    $form.AcceptButton = $okButton
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(225, 220)
    $cancelButton.Size = New-Object System.Drawing.Size(75, 25)
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($cancelButton)
    $form.CancelButton = $cancelButton
    
    $form.TopMost = $true
    $result = $form.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $textBox.Text.Trim()
    } else {
        return $null
    }
}
# Helper function to get word stem by removing common suffixes
function Get-WordStem {
    param([string]$word)
    $word = $word.ToLower()
    $suffixes = @('ing', 'ed', 'es', 's', 'y', 'er', 'est', 'ly')
    
    foreach ($suffix in $suffixes) {
        if ($word.EndsWith($suffix) -and ($word.Length - $suffix.Length -ge 3)) {
            # For 'ies' pattern replacing 'y'
            if ($suffix -eq 'es' -and $word.EndsWith('ies')) {
                return $word.Substring(0, $word.Length - 3) + 'y'
            }
            # For other suffixes
            return $word.Substring(0, $word.Length - $suffix.Length)
        }
    }
    return $word
}

# Helper function to get similarity score between two strings
function Get-StringSimilarity {
    param([string]$str1, [string]$str2)
    
    if ([string]::IsNullOrEmpty($str1) -or [string]::IsNullOrEmpty($str2)) {
        return 0
    }
    
    $str1 = $str1.ToLower()
    $str2 = $str2.ToLower()
    
    # Exact match gets highest score
    if ($str1 -eq $str2) {
        return 10
    }
    
    # Get word stems
    $stem1 = Get-WordStem $str1
    $stem2 = Get-WordStem $str2
    
    # Stem match gets high score
    if ($stem1 -eq $stem2) {
        return 8
    }
    
    # Check if one string contains the other
    if ($str1.Contains($str2) -or $str2.Contains($str1)) {
        return 6
    }
    
    # Check for partial matches with 3+ character segments
    $score = 0
    for ($i = 0; $i -lt ($str1.Length - 2); $i++) {
        $segment = $str1.Substring($i, [Math]::Min(3, $str1.Length - $i))
        if (($segment.Length -ge 3) -and $str2.Contains($segment)) {
            $score += 1
        }
    }
    
    return $score
}

# Function to match keywords in a prompt with spell components
function Match-SpellKeywords {
    param(
        [string]$prompt,
        [string]$componentType,
        [array]$components,
        [switch]$Debug = $false
    )
    
    # Normalize prompt to improve matching
    $normalizedPrompt = $prompt.ToLower()
    
    # Get the appropriate synonym dictionary based on component type
    $componentSynonyms = $synonyms[$componentType]
    
    # Track best matches with scores
    $bestMatch = $null
    $highestScore = 0
    
    if ($Debug) {
        Write-Host "Matching for component type: $componentType" -ForegroundColor Yellow
        Write-Host "Prompt: $normalizedPrompt" -ForegroundColor Gray
    }
    
    # Process each potential component
    foreach ($component in $components) {
        $score = 0
        $matchReason = ""

        # Check for exact match in the prompt
        if ($normalizedPrompt.Contains($component.ToLower())) {
            $score = 10
            $matchReason = "Exact match"
        }

        # Check synonyms for this component
        if ($componentSynonyms -and $componentSynonyms[$component]) {
            foreach ($synonym in $componentSynonyms[$component]) {
                # Skip empty synonyms
                if ([string]::IsNullOrWhiteSpace($synonym)) {
                    continue
                }

                # Check for exact synonym match
                if ($normalizedPrompt.Contains($synonym.ToLower())) {
                    $synonymScore = 9

                    # Longer synonym matches are more precise
                    if ($synonym.Length -gt 3) {
                        $synonymScore = 10
                    }

                    if ($synonymScore -gt $score) {
                        $score = $synonymScore
                        $matchReason = "Synonym match: $synonym"
                    }
                }
            }
        }

        # Special handling for Duration component type using regex patterns
        if ($componentType -eq "Duration" -and $score -eq 0) {
            # Debug info for duration patterns
            if ($Debug) {
                Write-Host "  Checking duration regex patterns for $component..."
            }
            
            # Check if we have patterns for this component
            if ($timingPatterns.duration_patterns.$component -and 
                $timingPatterns.duration_patterns.$component.regex_patterns) {
                
                $regexPatterns = $timingPatterns.duration_patterns.$component.regex_patterns
                
                foreach ($pattern in $regexPatterns) {
                    if ($normalizedPrompt -match $pattern) {
                        $score = 10
                        $matchReason = "Duration pattern match: $pattern"
                        
                        if ($Debug) {
                            Write-Host "  Found regex match for '$component' with pattern: $pattern" -ForegroundColor Cyan
                        }
                        break
                    }
                }
            }
        }

        # If no exact, synonym, or pattern match, try stem matching
        if ($score -eq 0) {
            # Get stem of component
            $componentStem = Get-WordStem -word $component
            
            # Check if component stem exists in prompt
            if ($componentStem.Length -ge 3 -and $normalizedPrompt.Contains($componentStem)) {
                $score = 6
                $matchReason = "Stem match: $componentStem"
            }

            # Try synonym stems if component stem didn't match
            if ($score -eq 0 -and $componentSynonyms -and $componentSynonyms[$component]) {
                foreach ($synonym in $componentSynonyms[$component]) {
                    if ([string]::IsNullOrWhiteSpace($synonym)) {
                        continue
                    }

                    $synonymStem = Get-WordStem -word $synonym
                    if ($synonymStem.Length -ge 3 -and $normalizedPrompt.Contains($synonymStem)) {
                        $score = 5
                        $matchReason = "Synonym stem match: $synonymStem (from $synonym)"
                        break
                    }
                }
            }
        }

        # If still no match, try string similarity as a fallback
        if ($score -eq 0) {
            # Check component similarity
            $similarityScore = Get-StringSimilarity -str1 $component -str2 $normalizedPrompt

            if ($similarityScore -gt 0) {
                $score = $similarityScore
                $matchReason = "Similarity match with score $similarityScore"
            }

            # Check synonym similarity
            if ($componentSynonyms -and $componentSynonyms[$component]) {
                foreach ($synonym in $componentSynonyms[$component]) {
                    if ([string]::IsNullOrWhiteSpace($synonym)) {
                        continue
                    }

                    $synSimilarityScore = Get-StringSimilarity -str1 $synonym -str2 $normalizedPrompt
                    if ($synSimilarityScore -gt $score) {
                        $score = $synSimilarityScore
                        $matchReason = "Synonym similarity match: $synonym with score $synSimilarityScore"
                    }
                }
            }
        }

        # Update best match if current component has higher score
        if ($score -gt $highestScore) {
            $highestScore = $score
            $bestMatch = $component
            
            if ($Debug) {
                Write-Host "  New best match: $component with score $score ($matchReason)" -ForegroundColor Cyan
            }
        }
    }

    # Return the best match if the score is high enough
    # A minimum threshold prevents weak matches
    if ($highestScore -ge 3) {
        return $bestMatch
    }
    else {
        if ($Debug) {
            Write-Host "  No good matches found (highest score: $highestScore)" -ForegroundColor Red
        }
        return $null
    }
}

# Function to analyze a text prompt and extract spell components
function Get-SpellComponentsFromPrompt {
    param(
        [string]$prompt
    )
    
    # Initialize spell components with null values
    $components = @{
        "Effect" = $null
        "Element" = $null
        "Level" = $null
        "Duration" = $null
        "Range" = $null
    }
    
    # Get available components from the spell data
    $effects = $spellData.spoken_spell_table.effect_prefix.PSObject.Properties.Name
    $elements = $spellData.spoken_spell_table.element_prefix.PSObject.Properties.Name
    $durations = $spellData.spoken_spell_table.duration_modifier.PSObject.Properties.Name
    $ranges = $spellData.spoken_spell_table.range_suffix.PSObject.Properties.Name
    
    # Normalize prompt to improve matching
    $normalizedPrompt = $prompt.ToLower()
    
    # Use the Match-SpellKeywords function for components
    if (-not $components.Effect) {
        $components.Effect = Match-SpellKeywords -prompt $prompt -componentType "Effect" -components $effects -Debug
        if ($Debug) {
            Write-Host "Effect component result: $($components.Effect)" -ForegroundColor Cyan
        }
    }

    if (-not $components.Element) {
        $components.Element = Match-SpellKeywords -prompt $prompt -componentType "Element" -components $elements -Debug
        if ($Debug) {
            Write-Host "Element component result: $($components.Element)" -ForegroundColor Cyan
        }
    }

    if (-not $components.Duration) {
        $components.Duration = Match-SpellKeywords -prompt $prompt -componentType "Duration" -components $durations -Debug
    }

    if (-not $components.Range) {
        $components.Range = Match-SpellKeywords -prompt $prompt -componentType "Range" -components $ranges -Debug
    }
    
    # Try to identify level from numeric values in the prompt
    $levelPattern = "level\s+(\d+)|(\d+)\s*(?:st|nd|rd|th)?\s+level"
    if ($prompt -match $levelPattern) {
        $levelValue = if ($matches[1]) { $matches[1] } else { $matches[2] }
        # Ensure level is within valid range (1-10)
        $levelInt = [int]$levelValue
        if ($levelInt -ge 1 -and $levelInt -le 10) {
            $components.Level = "$levelInt"
        }
    }
    
    # For any missing components, make an additional attempt with Match-SpellKeywords
    if (-not $components.Effect) {
        # Try again with a default fallback
        $components.Effect = Match-SpellKeywords -prompt $prompt -componentType "Effect" -components $effects -Debug
        
        # If we still don't have a component, default to Creation
        if (-not $components.Effect) {
            $components.Effect = "Creation"
        }
    }
    
    # Try to infer element if not already set
    if (-not $components.Element) {
        # Try again with a default fallback
        $components.Element = Match-SpellKeywords -prompt $prompt -componentType "Element" -components $elements -Debug
        
        # If we still don't have a component, default to Moon
        if (-not $components.Element) {
            $components.Element = "Moon"
        }
    }
    
    # Try to determine duration if not set
    if (-not $components.Duration) {
        # Try again with a default fallback
        $components.Duration = Match-SpellKeywords -prompt $prompt -componentType "Duration" -components $durations -Debug
        
        # If we still don't have a component, set appropriate defaults based on effect type
        if (-not $components.Duration) {
            # Default duration based on effect type
            switch ($components.Effect) {
                "Creation" { $components.Duration = "10_minute" }
                "Damage" { $components.Duration = "Instant" }
                "Shield" { $components.Duration = "5_minute" }
                "Heal" { $components.Duration = "Instant" }
                default { $components.Duration = "1_minute" }
            }
        }
    }
    
    # Try to determine range if not set
    if (-not $components.Range) {
        # Try again with a default fallback
        $components.Range = Match-SpellKeywords -prompt $prompt -componentType "Range" -components $ranges -Debug
        
        # If we still don't have a component, default to 30ft
        if (-not $components.Range) {
            $components.Range = "30ft"
        }
    }
    
    # If level not set, default to level 1
    if (-not $components.Level) {
        $components.Level = "1"
    }
    
    # Debug output for final component selections
    # Debug output for final component selections
    Write-Host "Final component selections:" -ForegroundColor Cyan
    Write-Host "  Effect: $($components.Effect)" -ForegroundColor Green
    Write-Host "  Element: $($components.Element)" -ForegroundColor Green
    Write-Host "  Level: $($components.Level)" -ForegroundColor Green
    Write-Host "  Duration: $(Format-DurationText $($components.Duration))" -ForegroundColor Green
    Write-Host "  Range: $($components.Range)" -ForegroundColor Green
    return $components
}

# Function to calculate spell efficiency based on caster bloodline and spell element
function Get-SpellEfficiency {
    param(
        [string]$bloodline,
        [string]$element
    )
    
    # Default to neutral if no match is found
    $efficiency = "Neutral 50%"
    $percentage = 50
    
    # Handle Sun bloodline (neutral with everything)
    if ($bloodline -eq "Sun") {
        return @{
            "EfficiencyLabel" = "Neutral 50%"
            "EfficiencyPercentage" = 50
        }
    }
    
    # Check if bloodline matches element exactly (Best 100%)
    if ($bloodline -eq $element) {
        return @{
            "EfficiencyLabel" = "Best 100%"
            "EfficiencyPercentage" = 100
        }
    }
    
    # Check each efficiency level for the bloodline
    foreach ($level in @("Best 80%", "Good 60%", "Moderate 40%", "Weak 20%", "Neutral 50%")) {
        $elements = $spellData.bloodline_affinities.$bloodline.$level
        
        # Special case for Sun element
        if ($element -eq "Sun") {
            return @{
                "EfficiencyLabel" = "Neutral 50%"
                "EfficiencyPercentage" = 50
            }
        }
        
        # Check if the element is in this efficiency level
        if ($elements -contains $element) {
            $efficiency = $level
            
            # Extract percentage from the efficiency level
            if ($level -match "(\d+)%") {
                $percentage = [int]$matches[1]
            }
            
            break
        }
    }
    
    return @{
        "EfficiencyLabel" = $efficiency
        "EfficiencyPercentage" = $percentage
    }
    }

    # Function to get a spell description based on effect, element, and other attributes
    # Function to get a spell description based on effect, element, and other attributes
    function Get-SpellDescription {
        param(
            [string]$effect,
            [string]$element,
            [string]$level,
            [string]$duration,
            [string]$range
        )
        
        try {
            # Format duration for display
            $formattedDuration = Format-DurationText -duration $duration
            
            # Use ElementMapper to find the best matching template element
            # Pass the array of template elements extracted from spell_descriptions.json
            $mappedElement = Get-MappedElement -SourceElement $element -TemplateElements $templateElements
            
            # Initialize an empty array for description templates
            $descriptionTemplates = @()
            
            # Check if the spellDescriptions object exists and has the new structure
            if ($null -ne $spellDescriptions -and 
                $spellDescriptions.PSObject.Properties.Name -contains "spoken_spell_table" -and
                $spellDescriptions.spoken_spell_table.PSObject.Properties.Name -contains "effect_prefix") {
                
                # Check if the effect exists in effect_prefix
                if ($spellDescriptions.spoken_spell_table.effect_prefix.PSObject.Properties.Name -contains $effect) {
                    $effectObj = $spellDescriptions.spoken_spell_table.effect_prefix.$effect
                    
                    # Check if this effect has element_prefix structure
                    if ($effectObj.PSObject.Properties.Name -contains "element_prefix") {
                        
                        # Check if the mapped element exists for this effect
                        if ($effectObj.element_prefix.PSObject.Properties.Name -contains $mappedElement) {
                            # Get templates array safely
                            $templates = $effectObj.element_prefix.$mappedElement
                            if ($templates -is [Array]) {
                                $descriptionTemplates = $templates
                            } elseif ($null -ne $templates) {
                                $descriptionTemplates = @($templates)
                            }
                        }
                        
                        # If no templates found for mapped element, try the original element
                        if ($null -eq $descriptionTemplates -or $descriptionTemplates.Count -eq 0) {
                            if ($effectObj.element_prefix.PSObject.Properties.Name -contains $element) {
                                # Get templates array safely
                                $templates = $effectObj.element_prefix.$element
                                if ($templates -is [Array]) {
                                    $descriptionTemplates = $templates
                                } elseif ($null -ne $templates) {
                                    $descriptionTemplates = @($templates)
                                }
                            }
                        }
                        
                        # If still no templates, try to get templates from "Generic" element
                        if ($null -eq $descriptionTemplates -or $descriptionTemplates.Count -eq 0) {
                            if ($effectObj.element_prefix.PSObject.Properties.Name -contains "Generic") {
                                # Get templates array safely
                                $templates = $effectObj.element_prefix.Generic
                                if ($templates -is [Array]) {
                                    $descriptionTemplates = $templates
                                } elseif ($null -ne $templates) {
                                    $descriptionTemplates = @($templates)
                                }
                            }
                        }
                    }
                }
                
                # Fallback to Generic effect if no templates found
                if ($null -eq $descriptionTemplates -or $descriptionTemplates.Count -eq 0) {
                    if ($spellDescriptions.spoken_spell_table.effect_prefix.PSObject.Properties.Name -contains "Generic") {
                        $genericEffect = $spellDescriptions.spoken_spell_table.effect_prefix.Generic
                        
                        if ($genericEffect.PSObject.Properties.Name -contains "element_prefix") {
                            # Try mapped element in Generic effect
                            if ($genericEffect.element_prefix.PSObject.Properties.Name -contains $mappedElement) {
                                $templates = $genericEffect.element_prefix.$mappedElement
                                if ($templates -is [Array]) {
                                    $descriptionTemplates = $templates
                                } elseif ($null -ne $templates) {
                                    $descriptionTemplates = @($templates)
                                }
                            }
                            
                            # Try original element in Generic effect
                            if (($null -eq $descriptionTemplates -or $descriptionTemplates.Count -eq 0) -and 
                                $genericEffect.element_prefix.PSObject.Properties.Name -contains $element) {
                                $templates = $genericEffect.element_prefix.$element
                                if ($templates -is [Array]) {
                                    $descriptionTemplates = $templates
                                } elseif ($null -ne $templates) {
                                    $descriptionTemplates = @($templates)
                                }
                            }
                            
                            # Try Generic element in Generic effect
                            if (($null -eq $descriptionTemplates -or $descriptionTemplates.Count -eq 0) -and 
                                $genericEffect.element_prefix.PSObject.Properties.Name -contains "Generic") {
                                $templates = $genericEffect.element_prefix.Generic
                                if ($templates -is [Array]) {
                                    $descriptionTemplates = $templates
                                } elseif ($null -ne $templates) {
                                    $descriptionTemplates = @($templates)
                                }
                            }
                        }
                    }
                }
            }
            
            # Filter out null or empty templates
            $validTemplates = @($descriptionTemplates | Where-Object { $_ -and $_.Length -gt 0 })
            
            # If we have valid description templates, select one randomly and format it
            if ($validTemplates.Count -gt 0) {
                $template = Get-Random -InputObject $validTemplates
                
                # Replace placeholders with actual values
                $description = $template -replace "{EFFECT}", $effect `
                                    -replace "{ELEMENT}", $element `
                                    -replace "{LEVEL}", $level `
                                    -replace "{DURATION}", $formattedDuration `
                                    -replace "{RANGE}", $range
                
                return $description
            }
            
            # Return default description if no templates found
            return "A $level level $element $effect spell with $formattedDuration duration and $range range."
        }
        catch {
            Write-Host "Error generating spell description: $_" -ForegroundColor Red
            return "A $level level $element $effect spell with $formattedDuration duration and $range range."
        }
    }
    # Function to display results in a dialog box
    function Show-ResultDialog {
    param(
        [string]$title,
        [string]$message
    )
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $title
    $form.Size = New-Object System.Drawing.Size(600, 580)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "Sizable"
    
    $textBox = New-Object System.Windows.Forms.RichTextBox
    $textBox.Location = New-Object System.Drawing.Point(10, 10)
    $textBox.Size = New-Object System.Drawing.Size(560, 480)
    $textBox.Multiline = $true
    $textBox.ScrollBars = "Vertical"
    $textBox.ReadOnly = $true
    $textBox.Font = New-Object System.Drawing.Font("Consolas", 10)
    $textBox.Text = $message
    $textBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $form.Controls.Add($textBox)
    
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Location = New-Object System.Drawing.Point(250, 500)
    $closeButton.Size = New-Object System.Drawing.Size(100, 30)
    $closeButton.Text = "OK"
    $closeButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $closeButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom
    $form.Controls.Add($closeButton)
    $form.AcceptButton = $closeButton
    
    $form.TopMost = $true
    $form.ShowDialog() | Out-Null
}

# Function to generate spells based on selected components
function Get-Spell {
    param(
        [int]$count = 1,
        [string]$effectPrefix = $null,
        [string]$bloodline = $null
    )
    
    $results = ""
    
    for ($i = 1; $i -le $count; $i++) {
        # Select components
        $effect = if ($effectPrefix) { $effectPrefix } else { Get-Random -InputObject $spellData.spoken_spell_table.effect_prefix.PSObject.Properties.Name }
        $element = Get-Random -InputObject $spellData.spoken_spell_table.element_prefix.PSObject.Properties.Name
        $level = Get-Random -InputObject $spellData.spoken_spell_table.level_modifier.PSObject.Properties.Name
        $duration = Get-Random -InputObject $spellData.spoken_spell_table.duration_modifier.PSObject.Properties.Name
        $range = Get-Random -InputObject $spellData.spoken_spell_table.range_suffix.PSObject.Properties.Name
        
        # Get the corresponding affixes from the data
        $effectPre = $spellData.spoken_spell_table.effect_prefix.$effect
        $elementPre = $spellData.spoken_spell_table.element_prefix.$element
        $levelMod = $spellData.spoken_spell_table.level_modifier.$level
        $durationMod = $spellData.spoken_spell_table.duration_modifier.$duration
        $rangeSuf = $spellData.spoken_spell_table.range_suffix.$range

        # Create the spell name
        $spellName = "$effectPre $elementPre $levelMod $durationMod $rangeSuf".Trim()
        
        # Get spell description
        # Get spell description
        $spellDescription = Get-SpellDescription -effect $effect -element $element -level $level -duration $duration -range $range
        
        # Format duration for display in case the description function failed
        $formattedDuration = Format-DurationText -duration $duration
        
        # If description generation failed, use a default
        if ([string]::IsNullOrEmpty($spellDescription)) {
            $spellDescription = "A $level level $element $effect spell with $formattedDuration duration and $range range."
        }
        # Calculate spell efficiency if bloodline is provided
        $efficiencyInfo = ""
        if ($bloodline) {
            $efficiency = Get-SpellEfficiency -bloodline $bloodline -element $element
            $efficiencyInfo = "  Caster Bloodline: $bloodline`r`n  Spell Efficiency: $($efficiency.EfficiencyLabel) ($($efficiency.EfficiencyPercentage)% mana efficiency)`r`n"
        }
        
        # Build the result string
        $results += "Spell #$i`r`n"
        $results += "  Spell Description: $spellDescription`r`n"
        $results += "  Spell Name: $spellName`r`n`r`n"
        $results += "  Effect: $effect`r`n"
        $results += "  Element: $element`r`n"
        $results += "  Level: $level`r`n"
        $results += "  Duration: $(Format-DurationText $duration)`r`n"
        $results += "  Range: $range`r`n"
        $results += $efficiencyInfo
        
    }
    
    return $results
}

# Function to generate a spell from components
function Get-SpellFromComponents {
    param(
        [hashtable]$components,
        [string]$bloodline = $null,
        [string]$description = ""
    )
    
    # Get the corresponding affixes from the data
    $effectPre = $spellData.spoken_spell_table.effect_prefix.($components.Effect)
    $elementPre = $spellData.spoken_spell_table.element_prefix.($components.Element)
    $levelMod = $spellData.spoken_spell_table.level_modifier.($components.Level)
    $durationMod = $spellData.spoken_spell_table.duration_modifier.($components.Duration)
    $rangeSuf = $spellData.spoken_spell_table.range_suffix.($components.Range)

    # Create the spell name
    $spellName = "$effectPre $elementPre $levelMod $durationMod $rangeSuf".Trim()
    
    # Get spell description
    $spellDescription = Get-SpellDescription -effect $components.Effect -element $components.Element -level $components.Level -duration $components.Duration -range $components.Range
    
    # Format duration for display in case the description function failed
    $formattedDuration = Format-DurationText -duration $components.Duration
    
    # If description generation failed, use a default
    if ([string]::IsNullOrEmpty($spellDescription)) {
        $spellDescription = "A $($components.Level) level $($components.Element) $($components.Effect) spell with $formattedDuration duration and $($components.Range) range."
    }
    # Calculate spell efficiency if bloodline is provided
    $efficiencyInfo = ""
    if ($bloodline) {
        $efficiency = Get-SpellEfficiency -bloodline $bloodline -element $components.Element
        $efficiencyInfo = "  Caster Bloodline: $bloodline`r`n  Spell Efficiency: $($efficiency.EfficiencyLabel) ($($efficiency.EfficiencyPercentage)% mana efficiency)`r`n"
    }
    
    # Build the result string
    $results = "Generated Spell`r`n"
    if ($description) {
        $results += "  Original Description: $description`r`n"
    }
    $results += "  Spell Description: $spellDescription`r`n"
    $results += "  Spell Name: $spellName`r`n`r`n"
    $results += "  Effect: $($components.Effect)`r`n"
    $results += "  Element: $($components.Element)`r`n"
    $results += "  Level: $($components.Level)`r`n"
    $results += "  Duration: $(Format-DurationText $($components.Duration))`r`n"
    $results += "  Range: $($components.Range)`r`n"
    $results += $efficiencyInfo
    return $results
}

# Add debug message to confirm script is running
Write-Host "DialogSpellGenerator started successfully" -ForegroundColor Green

# Main script execution
$welcomeMessage = @"
===== Magical Spell Generator =====

Welcome, Spellcaster!

This enchanted program will generate random spell names and attributes 
based on a carefully constructed magical language system.

Each mystical spell is composed of these arcane components:

Effect: What the spell does (Creation, Damage, Shield, etc.)
Element: The element or type of magic (Fire, Water, Song, etc.)
Level: The power level from 1-10
Duration: How long the spell's magic persists
Range: The distance the spell can reach

The spell name combines these components into a powerful magical incantation
that you can use in your adventures or creative works.

Each spellcaster has a magical bloodline that affects their efficiency
with different elemental magics. The program will calculate and display
your spell's efficiency based on your bloodline and the spell's element.

You can generate spells in three ways:
1. Random Generation: Create completely random spells
2. Specify Effect Type: Choose a specific effect for your spells
3. Describe Spell with Text: Enter a description and the program will
analyze it to identify spell components

Click OK to begin your magical journey.
"@

    # Display the welcome message
    Show-ResultDialog "Welcome to Spell Generator" $welcomeMessage

    # Main execution loop
    $continueGenerating = $true

    while ($continueGenerating) {
        # Starting new spell generation cycle
        # Ask how many spells to generate
        $spellCountSelected = $false
        $numSpells = $null

            while (-not $spellCountSelected) {
                $numSpells = Get-NumberInputDialog "Spell Generator" "How many spells would you like to generate?" 3
                
                if (-not $numSpells) {
                    $retry = Get-YesNoDialog "No Spell Count" "No number of spells was selected. Would you like to try again?"
                    if (-not $retry) {
                        $continueGenerating = $false
                        break
                    }
                } 
                else {
                    $spellCountSelected = $true
                }
            }
            
            # Skip further processing if user cancelled
            if (-not $continueGenerating) {
                break
            }
            
            # Ask user to select their bloodline
            $bloodlineSelected = $false
            $selectedBloodline = $null

            while (-not $bloodlineSelected) {
                $bloodlineOptions = $spellData.bloodline_affinities.PSObject.Properties.Name
                $selectedBloodline = Get-SelectionDialog "Select Bloodline" "Choose your magical bloodline:" $bloodlineOptions
                
                if (-not $selectedBloodline) {
                    $retry = Get-YesNoDialog "No Bloodline Selected" "No bloodline was selected. Would you like to try again?"
                    if (-not $retry) {
                        $continueGenerating = $false
                        break
                    }
                } 
                else {
                    $bloodlineSelected = $true
                }
            }
            
            # Skip further processing if user cancelled
            if (-not $continueGenerating) {
                break
            }
            
            # Ask user what method they would like to use
            $methodSelected = $false
            $method = $null

            while (-not $methodSelected) {
                $options = @("Random Generation", "Specify Effect Type", "Describe Spell with Text")
                $method = Get-SelectionDialog "Spell Generator" "How would you like to generate your spell?" $options
                
                if (-not $method) {
                    $retry = Get-YesNoDialog "No Method Selected" "No generation method was selected. Would you like to try again?"
                    if (-not $retry) {
                        $continueGenerating = $false
                        break
                    }
                } 
                else {
                    $methodSelected = $true
                }
            }
            
            # Skip further processing if user cancelled
            if (-not $continueGenerating) {
                break
            }

            $spellResults = ""
            
            # Generate spells based on selected method
            switch ($method) {
                    "Random Generation" {
                        # Generate spells without specific effect but with bloodline
                        $spellResults = Get-Spell -count $numSpells -bloodline $selectedBloodline
                    }
                    
                    "Specify Effect Type" {
                        # Let user select an effect from the list
                        $effects = $spellData.spoken_spell_table.effect_prefix.PSObject.Properties.Name
                        $selectedEffect = Get-SelectionDialog "Select Effect" "Choose the spell effect:" $effects
                        
                        if (-not $selectedEffect) {
                            Write-Host "No effect was selected" -ForegroundColor Yellow
                        } else {
                            # Generate spells with the selected effect and bloodline
                            $spellResults = Get-Spell -count $numSpells -effectPrefix $selectedEffect -bloodline $selectedBloodline
                        }
                    }
                    
                    "Describe Spell with Text" {
                        # Get text description from user
                        $spellDescription = Get-TextInputDialog "Describe Your Spell" "Enter a description of the spell you want to create:"
                        
                        if (-not $spellDescription) {
                            $spellResults = "No spell description was provided. Please try again."
                        } else {
                            # Extract components from the description
                            $components = Get-SpellComponentsFromPrompt -prompt $spellDescription
                            $spellResults = Get-SpellFromComponents -components $components -bloodline $selectedBloodline -description $spellDescription
                        }
                    }
                }
                
                # Display the spell results
                if ($spellResults) {
                    Show-ResultDialog "Your Magical Spells" $spellResults
                } else {
                    Write-Host "No spell results were generated" -ForegroundColor Yellow
                }
                # Ask if user wants to generate more spells
                # Ask if user wants to generate more spells
                $continueGenerating = Get-YesNoDialog "Continue?" "Would you like to generate more spells?"
            }
