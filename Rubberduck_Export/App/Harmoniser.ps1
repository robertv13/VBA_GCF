$keywords = @{
    '.address' = '.Address'
    '.borders' = '.Borders'
    '.cells' = '.Cells'
    '.column' = '.Column'
    '.count' = '.Count'
    '.font' = '.Font'
    '.interior' = '.Interior'
    '.list' = '.List'
    '.name' = '.Name'
    '.offset' = '.Offset'
    '.range' = '.Range'
    '.row' = '.Row'
    '.text' = '.Text'
    '.value' = '.Value'
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

# Supprimer tous les fichiers .frx dans le dossier d'export (et sous-dossiers) - 2025-06-25 @ 07:16
$frxFiles = Get-ChildItem -Recurse -Path $folder -Filter *.frx
if ($frxFiles.Count -gt 0) {
    foreach ($file in $frxFiles) {
        Remove-Item $file.FullName -Force
        Write-Host "🗑️ Supprimé : $($file.FullName)" -ForegroundColor DarkGray
    }
} else {
    Write-Host "✅ Aucun fichier .frx à supprimer" -ForegroundColor Green
}

pause
