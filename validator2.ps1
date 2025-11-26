param (
    [string]$folderPath
)

# verify "C:\ids_modelli\stefano\4.ifc" -ids "C:\ids_modelli\stefano\4.ids" -v "Normal"


# C:\XBIM-Validator2\ids-verify.exe verify |${CurrentItem.FullName}| -ids |${NomeIDS}| -v "Minimal" >|${Replaced}|


# Verificare che la cartella sia stata fornita come parametro
if (-not $folderPath) {
    Write-Host "Per favore, fornisci il percorso della cartella come parametro."
    exit
}

# Verificare che la cartella esista
if (Test-Path $folderPath) {
    # Ottenere tutti i file nella cartella
    $files = Get-ChildItem -Path $folderPath -File

    # Iterare su ogni file
    foreach ($file in $files) {
        # Definire il comando esterno che deve essere lanciato per ogni file
        # Sostituire 'ComandoEsterna' con il comando effettivo che desideri eseguire
        $comando = "C:\XBIM-Validator2\ids-verify.exe verify $($file.FullName) -ids \\pcromfirma\00_IDS_File\IDS_A24-PFTE_StralcioIX.ids -v ""Minimal"" >"
        
        # Esegui il comando esterno
        Write-Host "Eseguendo il comando per il file: $($file.Name)"
        Invoke-Expression $comando
    }
} else {
    Write-Host "La cartella specificata non esiste."
}