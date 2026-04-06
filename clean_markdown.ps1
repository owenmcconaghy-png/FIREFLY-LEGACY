$root = "C:\Users\owenm\OneDrive\Documents\firefly repo"

Get-ChildItem -Path $root -Recurse -Filter "*.md" | ForEach-Object {
    $content = Get-Content $_.FullName -Raw -Encoding UTF8
    
    # Skip files without HTML tags
    if ($content -notmatch '<p style=') { return }
    
    # Extract text from <p> tags
    $lines = @()
    $lines += "---"
    
    # Keep YAML frontmatter
    if ($content -match '(?s)^(---.*?---)') {
        $lines = @($matches[1])
        $content = $content -replace '(?s)^---.*?---\s*', ''
    }
    
    # Find all <p> blocks and extract text
    $pattern = '<p([^>]*)>(.*?)</p>'
    $matches_all = [regex]::Matches($content, $pattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)
    
    foreach ($m in $matches_all) {
        $style = $m.Groups[1].Value
        $text = $m.Groups[2].Value
        
        # Strip inner span tags
        $text = [regex]::Replace($text, '<[^>]+>', '')
        $text = $text.Trim()
        if (-not $text) { continue }
        
        # Map styles to markdown
        $isBold = $style -match 'font-weight:bold'
        $isCenter = $style -match 'text-align:center'
        $sizeMatch = [regex]::Match($style, 'font-size:(\d+)px')
        $size = if ($sizeMatch.Success) { [int]$sizeMatch.Groups[1].Value } else { 14 }
        $colorMatch = [regex]::Match($style, 'color:(#[0-9A-Fa-f]+)')
        $color = if ($colorMatch.Success) { $colorMatch.Groups[1].Value.ToUpper() } else { '' }
        
        if ($color -eq '#8B2500' -and $size -ge 40) {
            $lines += ""; $lines += "# $text"; $lines += ""
        } elseif ($color -eq '#C8960C' -and $size -ge 20) {
            $lines += ""; $lines += "## $text"; $lines += ""
        } elseif ($color -eq '#8B2500' -and $isBold -and $size -ge 14 -and -not $isCenter) {
            $lines += ""; $lines += "## $text"; $lines += ""
        } elseif ($isCenter -and $color -in @('#555555','#777777') -and $size -ge 13) {
            $lines += ""; $lines += "*$text*"; $lines += ""
        } elseif ($isBold -and -not $isCenter -and $size -le 16) {
            $lines += "**$text**"
        } elseif ($color -in @('#AAAAAA','#888888') -and $size -le 13) {
            $lines += "*$text*"
        } else {
            $lines += $text
        }
    }
    
    $clean = ($lines -join "`n") -replace "`n{3,}", "`n`n"
    Set-Content -Path $_.FullName -Value $clean.Trim() -Encoding UTF8
    Write-Host "Cleaned: $($_.Name)"
}

Write-Host "Done."