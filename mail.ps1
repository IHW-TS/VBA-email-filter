# Fonction pour lister les dossiers
function ListFolders($folder, $indent = 0) {
    $prefix = " " * $indent
    Write-Host "$prefix$($folder.Name)"
    foreach ($subFolder in $folder.Folders) {
        ListFolders $subFolder ($indent + 2)
    }
}

# Se connecter à Outlook
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$rootFolder = $namespace.Folders.Item("teoman.soykan@lcl.fr")

# Lister les dossiers
ListFolders $rootFolder
