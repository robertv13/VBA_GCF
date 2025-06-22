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

$files = Get-ChildItem -Recurse -Path $folder -Include *.bas, *.cls, *.frm

if ($files.Count -eq 0) {
    Write-Host "❌ Aucun fichier à traiter dans $folder" -ForegroundColor Yellow
} else {
    foreach ($file in $files) {
        Write-Host "🔧 Traitement : $($file.Name)"
        $content = Get-Content $file.FullName -Raw
        foreach ($pair in $keywords.GetEnumerator()) {
            $pattern = '(?<![\w])' + [regex]::Escape($pair.Key) + '(?![\w])'
            $content = [regex]::Replace($content, $pattern, $pair.Value)
        }
        Set-Content $file.FullName $content -Encoding UTF8
        Write-Host "✅ Fichier modifié : $($file.Name)"
    }
}
pause
