###############################################################
## PowerShell Zimbra Administration des comptes
## Version 0.1 - 06/06/2019
##
##  Exemples d'utilisation de la classe ZimbraAccount
##  avec une délégation d'authentification depuis un compte Admin
##
## Développée et testée sous Zimbra 8.8
## Nécessite PowerShell 5
## Nécessite PowerShell 6 pour ignorer les erreurs de certificat
##
## Auteur : Stéphane GOSNET
###############################################################

# Chargement classe zimbraAccount
. "$PSScriptRoot\zimbraAccount.ps1"
# Chargement classe zimbraAccount pour authentification déléguée
. "$PSScriptRoot\zimbraadmin.ps1"

# Instanciation zimbraAccount
$zimbra = [ZimbraAccount]::new("https://zimbra.gosnet.fr",8443,$true);
if(-not($zimbra.open("admin@zimbra.gosnet.fr","zimbra"))){
    Write-host "Erreur d'authentification !!"
}

###############################################################
# Instanciation zimbraAdmin et une delegation d'authentification sur zimbraAccount
$zimbraAdmin = [ZimbraAdmin]::new("https://zimbra.gosnet.fr",7071,$true);
if(-not($zimbraAdmin.open("admin@zimbra.gosnet.fr","zimbra"))){
    Write-host "Erreur d'authentification !!"
}
$delegatedToken = $zimbraAdmin.delegateAuth("stephane@zimbra.gosnet.fr",60)
$zimbra.delegateAuthAccount("stephane@zimbra.gosnet.fr",$delegatedToken)
###############################################################

# Template signature
$template = "<div>
<p style=`"margin: 0px;`"><span style=`"font-size: 10pt; font-family: 'calibri', 'sans-serif'; color: #c7373b;`">____________________________________________________________________________________</span></p>
<p style=`"margin: 0px;`"><span style=`"font-size: 10pt; font-family: 'trebuchet ms', sans-serif;`"><strong><span style=`"color: #c7373b;`">Pr&eacute;nom NOM<br /></span></strong></span></p>
<p style=`"margin: 0px;`"><span style=`"color: #999999;`"><strong><em><span style=`"font-size: 7.5pt; font-family: 'arial', 'sans-serif';`">Pour le respect de l'environnement, merci de n'imprimer ce mail qu'en cas de n&eacute;cessit&eacute;.</span></em></strong></span></p>
</div>"

# Ajoute une signature Html au compte
write-host "------------- Ajout Signature Template HTML"
$id_template = $zimbra.createAccountSignature("Template Html","text/html",$template)
Write-host $id_template

# Recherche signature d'un compte
write-host "------------- Recherche Signature Template HTML"
$signature = $zimbra.searchAccountSignature("Template Html")
write-host $signature.name ($signature.id) "type" $signature.type 

# Ajoute une signature text au compte
write-host "------------- Ajout Signature TXT"
$id_txt = $zimbra.createAccountSignature("Signature TXT","text/plain","Cordialement")
Write-host $id_txt

# Ajoute une signature Html au compte
write-host "------------- Ajout Signature HTML"
$id_html = $zimbra.createAccountSignature("Signature Html","text/html",$template)
Write-host $id_html

# Modifie une signature au compte
write-host "------------- Modifie Signature"
$signNew = $template.replace("Pr&eacute;nom NOM","Stephane GOSNET")
$result = $zimbra.modifyAccountSignature($id_html,"text/html",$signNew)
Write-host $result

# Liste signatures d'un compte
write-host "------------- Liste Signatures"
$signatures = $zimbra.getAccountSignatures()
foreach($signature in $signatures){
    Write-host $signature.name ($signature.id)
}

# Supprime la signature TXT du compte
write-host "------------- Supprime Signature TXT"
$result = $zimbra.deleteAccountSignature($id_txt)
Write-host $result

# Signature Html par défaut
write-host "------------- Signature par défaut"
write-host "------------------ Avant : Template défaut"
$attributs = @("zimbraPrefDefaultSignatureId","zimbraPrefForwardReplySignatureId")
$identites = $zimbra.getAccountIdentities($attributs)
write-host $identites

$zimbra.setAccountSignatureDefault($identites[0].id, $id_html)
$zimbra.setAccountSignatureReply($identites[0].id, $id_html)

write-host "------------------ Après : Signature Html par défaut sur nouveau mail et réponse"
$attributs = @("zimbraPrefDefaultSignatureId","zimbraPrefForwardReplySignatureId")
$identites = $zimbra.getAccountIdentities($attributs)
write-host $identites

# Suppression Signature Reply
write-host "------------- Suppresion de la signature Reply"
write-host "------------------ Avant : Html"
$attributs = @("zimbraPrefDefaultSignatureId","zimbraPrefForwardReplySignatureId")
$identites = $zimbra.getAccountIdentities($attributs)
write-host $identites

$zimbra.removeAccountSignatureReply($identites[0].id)

write-host "------------------ Après : Sans signature"
$attributs = @("zimbraPrefDefaultSignatureId","zimbraPrefForwardReplySignatureId")
$identites = $zimbra.getAccountIdentities($attributs)
write-host $identites

###############################################################
# Libération de la délégation : ici l'utilisateur connecté à zimbraAccount n'agit plus que sur son compte
$zimbra.releaseDelegateAuthAccount()
###############################################################

# Fermeture des sessions
echo "---------- Fermeture session ---------------------"
if($zimbraAdmin.close()){
    write-host "Session Zimbra Admin terminée"
}else{
    write-host "Erreur de fermeture de la session"
}
if($zimbra.close()){
    write-host "Session Zimbra terminée"
}else{
    write-host "Erreur de fermeture de la session"
}
