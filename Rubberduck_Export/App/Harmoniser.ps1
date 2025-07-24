# 📘 Dictionnaire typographique — casse normalisée
$keywords = @{
    '.address' = '.Address'
    '.borders' = '.Borders'
    '.cells' = '.Cells'
    '.column' = '.Column'
    '.count' = '.Count'
    '.font' = '.Font'
    '.interior' = '.Interior'
    '.name' = '.Name'
    '.offset' = '.Offset'
    '.range' = '.Range'
    '.row' = '.Row'
    '.text' = '.Text'
    '.value' = '.Value'
    '.worksheetfunction' = '.WorksheetFunction'
    '.goto' = '.GoTo'        # ✅ correction ciblée
    '.listbox' = '.ListBox'  # ✅ éléments visuels dans les .frm
}

# 📁 Dossier racine
$folder = $PSScriptRoot
$journal = "$folder\journalNettoyage.txt"
Set-Content -Path $journal -Value "📘 Journal typographique — $(Get-Date)" -Encoding UTF8

# 🔍 Fichiers ciblés (comme dans ta version de base)
$files = Get-ChildItem -Recurse -Path $folder -Include *.bas, *.cls, *.doccls, *.frm |
    Where-Object { $_.Extension -ne ".frx" }

foreach ($file in $files) {
    Write-Host "🔧 Traitement : $($file.Name)"
    Add-Content -Path $journal -Value "`n🔧 Fichier : $($file.Name)"

    # 📥 Lecture
    $content = Get-Content $file.FullName -Raw

    # 🔄 Remplacement typographique insensible à la casse
    foreach ($pair in $keywords.GetEnumerator()) {
        $pattern = '(?i)(?<![\w])' + [regex]::Escape($pair.Key) + '(?![\w\d])'
        if ($content -match $pattern) {
            Add-Content -Path $journal -Value "   ✅ Remplacé : $($pair.Key) ➜ $($pair.Value)"
            $content = [regex]::Replace($content, $pattern, $pair.Value)
        }
    }

    # 🔒 Correction ciblée finale .Goto ➜ .GoTo (sécurisée)
    $patternGoto = '(?i)(?<![\w])\.goto(?![\w\d])'
    if ([regex]::Matches($content, $patternGoto).Count -gt 0) {
        $content = [regex]::Replace($content, $patternGoto, '.GoTo')
        Add-Content -Path $journal -Value "   🔄 Correction ciblée .Goto ➜ .GoTo"
    }

    # 📤 Réécriture en UTF-8 sans BOM
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    $writer = New-Object System.IO.StreamWriter($file.FullName, $false, $utf8NoBom)
    $writer.Write($content)
    $writer.Close()

    Write-Host "✅ Fichier harmonisé : $($file.Name)"
}

# 🗑️ Suppression des fichiers .frx inutiles
$frxFiles = Get-ChildItem -Recurse -Path $folder -Filter *.frx
foreach ($file in $frxFiles) {
    Remove-Item $file.FullName -Force
    Write-Host "🗑️ Supprimé : $($file.Name)" -ForegroundColor DarkGray
}

pause
