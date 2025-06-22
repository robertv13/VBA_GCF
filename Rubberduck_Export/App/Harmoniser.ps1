$keywords = @{
    '.row' = '.Row'
    '.column' = '.Column'
    '.cells' = '.Cells'
    '.range' = '.Range'
    '.value' = '.Value'
    '.text' = '.Text'
    '.address' = '.Address'
    '.offset' = '.Offset'
    '.count' = '.Count'
    '.name' = '.Name'
    '.font' = '.Font'
    '.interior' = '.Interior'
    '.borders' = '.Borders'
    '.worksheetfunction' = '.WorksheetFunction'
}

$folder = $PSScriptRoot

if (-Not (Test-Path $folder)) {
    Write-Host "❌ Dossier introuvable : $folder" -ForegroundColor Red
    pause
    exit
}

# Obtenir les bons fichiers, en excluant les .frx
$files = Get-ChildItem -Recurse -Path $folder -Include *.bas, *.cls, *.doccls, *.frm |
    Where-Object { $_.Extension -ne ".frx" }

if ($files.Count -eq 0) {
    Write-Host "❌ Aucun fichier à traiter dans $folder" -ForegroundColor Yellow
} else {
    foreach ($file in $files) {
        Write-Host "🔧 Traitement : $($file.Name)"

        # Lire le contenu
        $content = Get-Content $file.FullName -Raw

        # Appliquer les remplacements .row => .Row, etc.
        foreach ($pair in $keywords.GetEnumerator()) {
            $pattern = '(?<![\w])' + [regex]::Escape($pair.Key) + '(?![\w])'
            $content = [regex]::Replace($content, $pattern, $pair.Value)
        }

        # Supprimer les lignes visuelles sensibles dans les fichiers .frm uniquement
        if ($file.Extension -eq ".frm") {
            $contentLines = $content -split "`r?`n"
            $filteredLines = $contentLines | Where-Object {
                $_ -notmatch '^\s*(ClientHeight|ClientWidth|StartUpPosition|Left|Top|Zoom|ScrollBars|WindowState)\s*='
            }
            $content = $filteredLines -join "`r`n"
        }

        # Réécrire le fichier
        Set-Content $file.FullName $content -Encoding UTF8
        Write-Host "✅ Fichier modifié : $($file.Name)"
    }
}

pause
