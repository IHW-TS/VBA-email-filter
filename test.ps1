# Fonction pour convertir un texte en UTF-8
function ConvertTo-UTF8($text) {
    return [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::Default.GetBytes($text))
}

# Fonction pour trouver un dossier par son nom
function Get-Folder($parentFolder, $folderName) {
    $utf8FolderName = ConvertTo-UTF8 $folderName
    foreach ($folder in $parentFolder.Folders) {
        if (ConvertTo-UTF8 $folder.Name -eq $utf8FolderName) {
            return $folder
        }
    }
    return $null
}

# Fonction pour trouver un dossier de manière récursive
function Find-Folder($parentFolder, $folderName) {
    $utf8FolderName = ConvertTo-UTF8 $folderName
    foreach ($folder in $parentFolder.Folders) {
        if (ConvertTo-UTF8 $folder.Name -eq $utf8FolderName) {
            return $folder
        } else {
            $foundFolder = Find-Folder $folder $folderName
            if ($foundFolder -ne $null) {
                return $foundFolder
            }
        }
    }
    return $null
}

# Ouvrir le fichier Excel
$excelPath = "$env:USERPROFILE\\Desktop\\aaaa.xlsx" # Nom de votre fichier Excel à placer sur votre bureau
$excel = New-Object -ComObject Excel.Application

try {
    $workbook = $excel.Workbooks.Open($excelPath)
    Write-Host "Fichier ouvert avec succès"
} catch {
    Write-Host "Erreur lors de l'ouverture du fichier Excel. Vérifiez le chemin et les permissions."
    exit
}

$sheet = $workbook.Sheets.Item("blanc") # Nom de votre feuille sur Excel

# Lire les numéros de dossier depuis la colonne A
$dossierNumbers = @()
$row = 1
while ($sheet.Cells.Item($row, 1).Value() -ne $null) {
    $dossierNumbers += $sheet.Cells.Item($row, 1).Value()
    $row++
}

# Afficher les numéros de dossier lus depuis Excel
Write-Host "Numéros de dossier lus depuis Excel :"
foreach ($num in $dossierNumbers) {
    Write-Host $num
}

# Fermer le fichier Excel
$workbook.Close($false)
$excel.Quit()

# Se connecter à Outlook
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$rootFolder = $namespace.Folders.Item("Votre boite mail générique") # Votre boite mail générique

# Trouver le dossier "Boîte de réception" de manière récursive
$inbox = Find-Folder -parentFolder $rootFolder -folderName "Boîte de réception"
if ($inbox -eq $null) {
    Write-Host "Erreur : Dossier 'Boîte de réception' introuvable."
    exit
}

# Trouver le dossier "test" au niveau supérieur
$testFolder = Get-Folder -parentFolder $rootFolder -folderName "test" # Le nom de votre sous dossier 
if ($testFolder -eq $null) {
    Write-Host "Erreur : Dossier 'test' introuvable."
    exit
}

Write-Host "Dossier 'test' trouvé avec succès"

# Parcourir les emails et les déplacer
foreach ($mail in $inbox.Items) {
    $subject = $mail.Subject
    Write-Host "Sujet de l'email : $subject"
    foreach ($dossierNumber in $dossierNumbers) {
        Write-Host "Comparaison avec le numéro de dossier : $dossierNumber"
        if ($subject -like "*$dossierNumber*") {
            Write-Host "Déplacement de l'email avec le sujet : $subject"
            $mail.Move($testFolder)
            break
        }
    }
}
