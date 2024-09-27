$excelPath = "$env:USERPROFILE\\Desktop\\aaaa.xlsx"
Write-Host "Chemin du fichier : $excelPath"

try {
    $excel = New-Object -ComObject Excel.Application
    $workbook = $excel.Workbooks.Open($excelPath)
    Write-Host "Fichier ouvert avec succès"
    $workbook.Close($false)
    $excel.Quit()
} catch {
    Write-Host "Erreur lors de l'ouverture du fichier Excel. Vérifiez le chemin et les permissions."
    Write-Host "Détails de l'erreur : $_"
}
