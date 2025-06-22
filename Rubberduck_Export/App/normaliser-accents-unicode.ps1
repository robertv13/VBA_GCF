param(
    [string]$Root = ""
)

if (-not (Test-Path $Root)) {
    Write-Host "❌ Le dossier spécifié est introuvable : $Root"
    exit
}

Add-Type -AssemblyName "System.Globalization"
$extensions = @("*.bas", "*.cls", "*.frm", "*.doccls")
$normalizer = [System.Text.NormalizationForm]::FormC

Get-ChildItem -Recurse -Path $Root -Include $extensions | ForEach-Object {
    $filePath = $_.FullName
    $changed = $false
    $lines = [System.IO.File]::ReadAllLines($filePath, [System.Text.Encoding]::UTF8)
    $newLines = foreach ($line in $lines) {
        $normalized = $line.Normalize($normalizer)
        if ($normalized -ne $line) {
            $changed = $true
        }
        $normalized
    }

    if ($changed) {
        Write-Host ("🎯 Normalisation : " + $filePath)
        [System.IO.File]::WriteAllLines($filePath, $newLines, [System.Text.Encoding]::UTF8)
    }
}
