###############################################################
## PowerShell Zimbra Administration
## Version 0.1 - 04/02/2019
##
##  Exemples d'utilisation de la classe ZimbraAdmin
##
## Développée et testée sous Zimbra 8.8
## Nécessite PowerShell 6 (pour ignorer les erreurs de certificat)
##
## Auteur : Stéphane GOSNET
###############################################################

# Chargement classe zimbraAdmin
. "$PSScriptRoot\zimbraAdmin.ps1"

# Instanciation zimbraAdmin et ouverture de Session
$zimbra = [ZimbraAdmin]::new("https://zimbra.gosnet.fr",7071,$true);
if(-not($zimbra.open("admin@zimbra.gosnet.fr","zimbra"))){
    Write-host "Erreur d'authentification !!"
}


# Recherche comptes par mail
$attributs = @("displayName","zimbraMailQuota","zimbraAccountStatus","zimbraMailAlias")
$comptes = $zimbra.getAccountsByMail("stephane@zimbra.gosnet.fr",$attributs)
foreach($compte in $comptes){
    echo "-------------------------------"
    Write-host "Compte :"$compte.name
    foreach($attribut in $compte.a){
        if($attribut.n -eq "displayName"){ write-host "Nom = "$attribut.'#text'}
        if($attribut.n -eq "zimbraMailDeliveryAddress"){ write-host "Addresse principale = "$attribut.'#text'}
        if($attribut.n -eq "zimbraMailAlias"){ write-host "Alias = "$attribut.'#text'}
        if($attribut.n -eq "zimbraMailQuota"){ write-host "Quota = "$attribut.'#text'}
        if($attribut.n -eq "zimbraAccountStatus"){ write-host "Status = "$attribut.'#text'}
    }
}

# Recherche comptes par Société
$attributs = @("displayName","company","title")
$comptes = $zimbra.getAccountsByCompany("GOSNET",$attributs)
foreach($compte in $comptes){
    echo "-------------------------------"
    Write-host "Compte :"$compte.name
    foreach($attribut in $compte.a){
        if($attribut.n -eq "displayName"){ write-host "Nom = "$attribut.'#text'}
        if($attribut.n -eq "company"){ write-host "Société = "$attribut.'#text'}
        if($attribut.n -eq "title"){ write-host "Fonction = "$attribut.'#text'}
    }
}

# Recherche comptes par fonction
$attributs = @("displayName","company","title")
$comptes = $zimbra.getAccountsByTitle("BOSS",$attributs)
foreach($compte in $comptes){
    echo "-------------------------------"
    Write-host "Compte :"$compte.name
    foreach($attribut in $compte.a){
        if($attribut.n -eq "displayName"){ write-host "Nom = "$attribut.'#text'}
        if($attribut.n -eq "company"){ write-host "Société = "$attribut.'#text'}
        if($attribut.n -eq "title"){ write-host "Fonction = "$attribut.'#text'}
    }
}

# Vérifie la présence du adresse
if($zimbra.testMailExist("spam@zimbra.gosnet.fr")){Write-host "L'adresse existe"}else{Write-Host "L'adresse n'existe pas"}


# Fermeture session
$zimbra.close()
