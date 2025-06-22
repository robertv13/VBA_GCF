param(
    [string]$Root = ""
)

if (-not (Test-Path $Root)) {
    Write-Host "❌ Le dossier spécifié est introuvable : $Root"
    exit
}

$extensions = @("*.bas", "*.cls", "*.frm", "*.doccls")
$propertyMap = @{}

# Étape 1 : Repérer les noms de propriétés Property Get/Let/Set
Get-ChildItem -Recurse -Path $Root -Include $extensions | ForEach-Object {
    $lines = [System.IO.File]::ReadAllLines($_.FullName, [System.Text.Encoding]::UTF8)
    foreach ($line in $lines) {
        if ($line -match "^\s*Property\s+(Get|Let|Set)\s+([a-zA-Z_][a-zA-Z0-9_]*)") {
            $name = $matches[2]
            $key = $name.ToLower()
            if (-not $propertyMap.ContainsKey($key)) {
                $propertyMap[$key] = $name
            }
        }
    }
}

# Étape 2 : Harmoniser tous les usages trouvés
Get-ChildItem -Recurse -Path $Root -Include $extensions | ForEach-Object {
    $filePath = $_.FullName
    Write-Host ("Traitement : " + $filePath)
    $lines = [System.IO.File]::ReadAllLines($filePath, [System.Text.Encoding]::UTF8)
    $newLines = foreach ($line in $lines) {
        $current = $line
        foreach ($key in $propertyMap.Keys) {
            $ref = [regex]::Escape($key)
            $pattern = '\b' + $ref + '\b'
            $current = [regex]::Replace($current, $pattern, $propertyMap[$key], [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        }
        $current
    }
    [System.IO.File]::WriteAllLines($filePath, $newLines, [System.Text.Encoding]::UTF8)
}
