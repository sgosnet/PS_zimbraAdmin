﻿###############################################################
## PowerShell Zimbra Administration
## Version 0.1 - 07/02/2019
##
##  Exemples d'utilisation de la classe ZimbraAdmin
##
## Développée et testée sous Zimbra 8.8
## Nécessite PowerShell 5
## Nécessite PowerShell 6 pour ignorer les erreurs de certificat
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
write-host "------------- Comptes par mail"
$attributs = @("displayName","name","zimbraMailQuota","zimbraAccountStatus","zimbraMailAlias")
$comptes = $zimbra.getAccountsByMail("*@zimbra.gosnet.fr",$attributs)
foreach($compte in $comptes){
    write-host ($compte)
}


# Recherche comptes par Société
write-host "------------- Comptes par société"
$attributs = @("displayName","company","title")
$comptes = $zimbra.getAccountsByCompany("GOSNET",$attributs)
foreach($compte in $comptes){
    write-host ($compte)
}

# Recherche comptes par fonction
write-host "------------- Comptes par fonction"
$attributs = @("displayName","company","title")
$comptes = $zimbra.getAccountsByTitle("BOSS",$attributs)
foreach($compte in $comptes){
    write-host ($compte)
}

# Vérifie la présence du adresse
write-host "------------- Test adresse"
if($zimbra.testMailExist("spam@zimbra.gosnet.fr")){Write-host "L'adresse existe"}else{Write-Host "L'adresse n'existe pas"}

# Recherche des comptes verouillés
write-host "------------- Comptes verouillés"
$attributs = @("displayName","zimbraAccountStatus")
$comptes = $zimbra.getAccountsLocked($attributs)
foreach($compte in $comptes){
    write-host ($compte)
}


# Fermeture session
echo "---------- Fermeture session ---------------------"
if($zimbra.close()){
    write-host "Session Zimbra Admin terminée"
}else{
    write-host "Erreur de fermeture de la session"
}
