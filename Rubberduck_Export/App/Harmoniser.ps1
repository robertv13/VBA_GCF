$keywords = @{
    '.address' = '.Address'
    '.borders' = '.Borders'
    '.cells' = '.Cells'
    '.column' = '.Column'
    '.count' = '.Count'
    '.font' = '.Font'
    '.goto' = '.GoTo'
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
$journal = "$folder\journalNettoyage.txt"
Set-Content -Path $journal -Value "📘 Journal de nettoyage typographique — $(Get-Date)" -Encoding UTF8

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
        Add-Content -Path $journal -Value "`n🔧 $($file.Name)"
        Write-Host "🔧 Traitement : $($file.Name)"
        
        # Lire le contenu
        $content = Get-Content $file.FullName -Raw
        
        # Appliquer les remplacements typographiques
        foreach ($pair in $keywords.GetEnumerator()) {
            $pattern = '(?<![\w])' + [regex]::Escape($pair.Key) + '(?![\w])'
            if ($content -match $pattern) {
                Add-Content -Path $journal -Value "   Remplacé : $($pair.Key) ➜ $($pair.Value)"
            }
            $content = [regex]::Replace($content, $pattern, $pair.Value)
        }

        # Nettoyer visuellement les fichiers .frm
        if ($file.Extension -eq ".frm") {
            $contentLines = $content -split "`r?`n"
            $filteredLines = $contentLines | Where-Object {
                ($_ -match '^\s*$') -or
                ($_ -notmatch '^\s*(ClientHeight|ClientWidth|StartUpPosition|Left|Top|Zoom|ScrollBars|WindowState)\s*=')
            }
            $content = $filteredLines -join "`r`n"
        }

        # Réécriture en UTF-8 sans BOM
        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        $writer = New-Object System.IO.StreamWriter($file.FullName, $false, $utf8NoBom)
        $writer.Write($content)
        $writer.Close()

        Write-Host "✅ Fichier modifié : $($file.Name)"
    }
}

# Suppression des fichiers .frx inutiles
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
