<#
.SYNOPSIS
Maps elements from spell_data.json to their closest equivalents in spell_descriptions.json templates.

.DESCRIPTION
This script provides mapping functionality between element systems to ensure compatibility
between the spell data and the available templates.
#>

# No JSON loading in this script - will be done by the main script

# Create a direct mapping table for known equivalents
$directMappings = @{
    "Wind" = "Air"
    "Moon" = "Light"
    "Song" = "Sound"
    "Love" = "Mind"
    "Protection" = "Shield"
}

# Simple string similarity function based on common characters
function Get-SimpleStringSimilarity {
    param (
        [string]$str1,
        [string]$str2
    )
    
    $str1 = $str1.ToLower()
    $str2 = $str2.ToLower()
    
    # If strings are identical, return 1.0
    if ($str1 -eq $str2) {
        return 1.0
    }
    
    # Count common characters
    $commonChars = 0
    foreach ($char in $str1.ToCharArray()) {
        if ($str2.Contains($char)) {
            $commonChars++
        }
    }
    
    # Calculate similarity score based on common chars vs total length
    $maxLength = [Math]::Max($str1.Length, $str2.Length)
    $similarity = $commonChars / $maxLength
    
    # Boost score if strings start with the same letter
    if ($str1[0] -eq $str2[0]) {
        $similarity += 0.1
        # Cap at 1.0
        if ($similarity -gt 1.0) { $similarity = 1.0 }
    }
    
    return $similarity
}

function Get-BestElementMatch {
    param (
        [string]$sourceElement,
        [string[]]$templateElements
    )
    
    # If element exists directly in templates, use it
    if ($sourceElement -in $templateElements) {
        return $sourceElement
    }
    
    # Check direct mapping table
    if ($directMappings.ContainsKey($sourceElement)) {
        return $directMappings[$sourceElement]
    }
    
    # Find closest match based on string similarity
    $bestMatch = "Any"
    $bestScore = 0.3  # Minimum threshold for similarity
    
    foreach ($templateElement in $templateElements) {
        $similarity = Get-SimpleStringSimilarity -str1 $sourceElement -str2 $templateElement
        if ($similarity -gt $bestScore) {
            $bestScore = $similarity
            $bestMatch = $templateElement
        }
    }
    
    # Return the best match (will be "Any" if no good match found)
    return $bestMatch
}

function Get-MappedElement {
    param (
        [string]$sourceElement,
        [string[]]$templateElements,
        [string]$targetSystem = "template"
    )
    
    # This function is a wrapper around Get-BestElementMatch for compatibility
    return Get-BestElementMatch -sourceElement $sourceElement -templateElements $templateElements
}

# Generate a report of all mappings
function Get-ElementMappingReport {
    param (
        [string[]]$spellDataElements,
        [string[]]$templateElements
    )
    
    $report = @()
    
    foreach ($element in $spellDataElements) {
        $match = Get-BestElementMatch -sourceElement $element -templateElements $templateElements
        $similarityScore = if ($element -eq $match) { 
            1.0 
        } elseif ($directMappings.ContainsKey($element)) { 
            0.9 
        } else { 
            Get-SimpleStringSimilarity -str1 $element -str2 $match 
        }
        
        $report += [PSCustomObject]@{
            SpellDataElement = $element
            TemplateElement = $match
            SimilarityScore = [math]::Round($similarityScore, 2)
            IsDirectMatch = $element -eq $match
            IsPredefinedMapping = $directMappings.ContainsKey($element)
        }
    }
    
    return $report | Sort-Object -Property SpellDataElement
}

# No need to export functions when dot-sourcing

# If running as a script rather than imported as a module
if ($MyInvocation.InvocationName -ne ".") {
    # When running as a script, you must provide the elements to map
    param (
        [Parameter(Mandatory=$true)]
        [string[]]$SourceElements,
        
        [Parameter(Mandatory=$true)]
        [string[]]$TargetElements
    )
    
    Write-Host "Element Mapping Report:" -ForegroundColor Yellow
    Write-Host "-------------------------" -ForegroundColor Yellow
    $report = Get-ElementMappingReport -spellDataElements $SourceElements -templateElements $TargetElements
    $report | Format-Table -AutoSize
    
    Write-Host "`nMapping Summary:" -ForegroundColor Cyan
    $directMatches = ($report | Where-Object { $_.IsDirectMatch }).Count
    $predefinedMappings = ($report | Where-Object { -not $_.IsDirectMatch -and $_.IsPredefinedMapping }).Count
    $algorithmMappings = ($report | Where-Object { -not $_.IsDirectMatch -and -not $_.IsPredefinedMapping }).Count
    
    Write-Host "Direct matches: $directMatches" -ForegroundColor Green
    Write-Host "Predefined mappings: $predefinedMappings" -ForegroundColor Yellow
    Write-Host "Algorithm-determined mappings: $algorithmMappings" -ForegroundColor Magenta
}
