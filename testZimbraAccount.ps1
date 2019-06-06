###############################################################
## PowerShell Zimbra Administration des comptes
## Version 0.1 - 06/06/2019
##
##  Exemples d'utilisation de la classe ZimbraAccount
##  Méthodes liées aux comptes
##
## Développée et testée sous Zimbra 8.8
## Nécessite PowerShell 5
## Nécessite PowerShell 6 pour ignorer les erreurs de certificat
##
## Auteur : Stéphane GOSNET
###############################################################

# Chargement classe zimbraAccount
. "$PSScriptRoot\zimbraAccount.ps1"

# Instanciation zimbraAdmin et ouverture de Session
$zimbra = [ZimbraAccount]::new("https://zimbra.gosnet.fr",8443,$true);
if(-not($zimbra.open("admin@zimbra.gosnet.fr","zimbra"))){
    Write-host "Erreur d'authentification !!"
}

# Recherche signature d'un compte
write-host "------------- Signature"
$signature = $zimbra.getAccountSignature("admin@zimbra.gosnet.fr","Le Mans 2019")
Write-host $signature.content

# Ajoute une signature text au compte
write-host "------------- Ajout Signature TXT"
$signNew = $signature.content.replace("Pr&eacute;nom NOM","Stephane GOSNET")
$id = $zimbra.createAccountSignature("admin@zimbra.gosnet.fr","Le Mans TEXT","text/plain","Cordialement")
Write-host $id

# Ajoute une signature Html au compte
write-host "------------- Ajout Signature HTML"
$id = $zimbra.createAccountSignature("admin@zimbra.gosnet.fr","Le Mans 2020","text/html",$signature.content)
Write-host $id

# Modifie une signature au compte
write-host "------------- Modifie Signature"
$signNew = $signature.content.replace("Pr&eacute;nom NOM","Stephane GOSNET")
$result = $zimbra.modifyAccountSignature("admin@zimbra.gosnet.fr",$id,"text/html",$signNew)
Write-host $result


# Fermeture session
echo "---------- Fermeture session ---------------------"
if($zimbra.close()){
    write-host "Session Zimbra Admin terminée"
}else{
    write-host "Erreur de fermeture de la session"
}
