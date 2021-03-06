﻿###############################################################
## PowerShell Zimbra Administration
## Version 0.1 - 07/06/2019
##
## Classe ZimbraAccount permettant des appels SOAP aux API
## la modification des comptes Zimbra.
##
## Développée et testée sous Zimbra 8.8
## PowerShell 5 
## Nécessite PowerShell 6 pour ignorer les erreurs de certificat
##
## Auteur : Stéphane GOSNET
##
###############################################################

###############################################################
## URL des WSDL Zimbra
## https://zimbra-serveur:7071/service/wsdl/ZimbraService.wsdl

###############################################################
## Problèmes à résoudre :
##  - Encodage des accents dans les réponses XML

###############################################################
# Constructeur class
#  Instancie la classe ZimbraAccount
#  Paramètres :
#   - url : URL du serveur Zimbra (https://zimbra.domaine.com)
#   - $port : Port TCP (8443 par défaut)
#   - ignoreCertificateError : Flag pour ignorer les erreurs de certificat (CA inconnue, date expirée) - Uniquement sous PowerShell 6.1
class ZimbraAccount
{
hidden [String] $soapHeader1 = '<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:urn="urn:zimbra" xmlns:urn1="urn:zimbraAccount"><soapenv:Header><urn:Context>'
hidden [String] $soapHeader2 = '</urn:Context></soapenv:Header>'
hidden [String] $soapFoot = '</soapenv:Envelope>'
hidden [String] $urlSoap = $null
hidden [String] $token = $null
hidden [String] $loginName = $null
hidden [String] $loginToken = $null


    [String] $url
    [Int] $port = 8443
    [Boolean] $ignoreCertificateError
    [String] $name =""

    ZimbraAccount([String]$url,[Int]$port,[Switch]$ignoreCertificateError){

        $this.url = $url
        $this.port = $port
        $this.urlSoap = "$url"+":"+"$port/service/soap"
        $this.ignoreCertificateError = $ignoreCertificateError
    }

###############################################################
###############################################################
# Méthodes techniques
###############################################################
###############################################################

###############################################################
# Méthode request()
#  Envoie une requête XML/SOAP à Zimbra
#  Paramètre :
#   - xml : requête SOAP au format XML
#  Retourne une réponse XML/SOAP
    [xml] request([String] $context,[String] $soap){
        
        # Construction de la requete SOAP avec entête et Pied XML
        $request = [xml](($this.soapHeader1)+$context+($this.soapHeader2)+$soap+($this.soapFoot))

        # Appel WebService avec ou sans prise en compte des erreurs de certificat (Support avec Powershell 6.1)
        try{
            if($this.ignoreCertificateError){
                $response = Invoke-WebRequest $this.urlSoap -Body $request -ContentType 'text/xml' -Method Post -SkipCertificateCheck
            }else{
                $response = Invoke-WebRequest $this.urlSoap -Body $request -ContentType 'text/xml' -Method Post
            }
        }catch{
            $response = "<response>Error</response>"
        }

        return [xml]$response
    }

###############################################################
# Méthode xmlToObjects()
#  Convertit une réponse XML/SOAP De la GAL en tableau d'objets PowerShell
#  Paramètre :
#   - xml : réponse SOAP au format XML
#   - attributs : tableau des attributs SOAP à insérer dans l'objet 
#  Retourne un Object PowerShell
    hidden [Object] xmlToObjects([Object]$xml,[String[]]$properties,[String[]]$attributs){
        
        $objets = @()
        # Parcours chaque résultat de la requete SOAP
        foreach($result in $xml){
            $objet = New-Object PsObject
            foreach($propertie in $properties){
                $objet | Add-member -Name "$propertie" -MemberType NoteProperty -Value ($result.($propertie))
            }
            $attrCnt=@{}
            # Comptage pour chaque attributs
            foreach($ligne in $result.a){
                foreach($attribut in $attributs){
                    if($ligne.name -eq $attribut){
                        $attrCnt[$attribut] += 1
                    }
                }
            }

            # Traitement des attributs
            $attrMulti=@{}
            foreach($ligne in $result.a){
                foreach($attribut in $attributs){
                    if($ligne.name -eq $attribut){
                        if($attrCnt[$attribut] -eq 1){
                            $objet | Add-member -Name $attribut -MemberType NoteProperty -Value ($ligne.'#text')
                        }else{
                            $attrMulti[$attribut] += 1
                            if($attrMulti[$attribut] -eq 1){
                                $objet | Add-member -Name $attribut -MemberType NoteProperty -Value @()
                            }
                            $objet.($attribut) += ($ligne.'#text')
                        }
                    }
                }
            }
            $objets += $objet
        }
        return $objets
    }

###############################################################
# Méthode open()
#  Ouvre une session Admin Zimbra avec un appel SOAP Auth-user
#  Initie une token Zimbra
#  Paramètres :
#   - login : Compte Admin
#   - password : Mot de passe Admin
#   - ignoreCertificateError : Flag pour ignorer les erreurs de certificat (CA inconnue, date expirée) - Uniquement sous PowerShell 6.1
#  Retourne $True ou $False si Erreur
    [Boolean] open([String]$login,[String]$password){
        # Requête SOAP d'Authentification
        $context = ""
        $request = "<soapenv:Body><urn1:AuthRequest><urn1:account by=`"adminName`">$login</urn1:account><urn1:password>$password</urn1:password></urn1:AuthRequest></soapenv:Body>"
        $response = $this.request($context,$request)

        # Enregistrement du Token et ajout dans l'entête SOAP
        if(($response.response) -ne "Error"){
            $this.loginName = $login
            $this.name = $login
            $this.loginToken = $response.Envelope.Body.Authresponse.authToken
            $this.token = $response.Envelope.Body.Authresponse.authToken
            $this.soapHeader1 = '<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:urn="urn:zimbra" xmlns:urn1="urn:zimbraAccount"><soapenv:Header><urn:context><urn:authToken>'+($this.token)+'</urn:authToken>'
            $this.soapHeader2 = '</urn:context></soapenv:Header>'
            return $True
        }else{
            return $False
        }
    }

###############################################################
# Méthode delegateAuthAccount()
#  Délègue une authentification sur un compte 
#  Initie un delegated token Zimbra sur un compte
#  Paramètres :
#   - $name : Nom du compte
    [Boolean] delegateAuthAccount([String]$name,[String]$delegatedToken){
        $this.name = $name
        $this.token = $delegatedToken
        $this.soapHeader1 = '<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:urn="urn:zimbra" xmlns:urn1="urn:zimbraAccount"><soapenv:Header><urn:context><urn:authToken>'+($this.token)+'</urn:authToken>'
        $this.soapHeader2 = '</urn:context></soapenv:Header>'
        return $True
    }

###############################################################
# Méthode releaseDelegateAuthAccount()
#  Libère la délèguation d'authentification sur un compte 
    [Boolean] releaseDelegateAuthAccount(){
        $this.name = $this.loginName
        $this.token = $this.loginToken
        $this.soapHeader1 = '<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:urn="urn:zimbra" xmlns:urn1="urn:zimbraAccount"><soapenv:Header><urn:context><urn:authToken>'+($this.token)+'</urn:authToken>'
        $this.soapHeader2 = '</urn:context></soapenv:Header>'
        return $True
    }


###############################################################
# Méthode close()
#  Ferme la session Admin Zimbra avec un appel SOAP Auth-user
#  Supprime le token Zimbra
#  $True ou $False si Erreur
    [Boolean] close(){
        $this.token = $null
        return $true
    }




###############################################################
###############################################################
# Méthodes liées aux Signatures de comptes
###############################################################
###############################################################

###############################################################
# Méthode getAccountSignatures()
#  Renvoie les signatures d'un compte
#  Paramètres: aucun
#  Retourne un tableau associatif des signatures (id, name, content)
    [Object]getAccountSignatures(){
        # Construction Requête SOAP
        $context = "<urn:account by=`"name`">"+($this.name)+"</urn:account>"
        $request = "<soapenv:Body><urn1:GetSignaturesRequest/>"
        $request += "</soapenv:Body>"
        $response = $this.request($context,$request)
        # Test de la réponse SOAP et renvoie d'un tableau des signatures
        if(($response.response) -ne "Error"){
            $signatures = $response.Envelope.Body.GetSignaturesResponse
            ([String]$signatures)
            $reponse = @()
            foreach($signature in $signatures.signature){
                $reponse += @{"id"=$signature.id;"name"=$signature.name;"type"=$signature.content.type;"content"=$signature.content.'#text'}
            }
            return $reponse
        }else{
            return $False
        }
    }

###############################################################
# Méthode searchAccountSignature()
#  Cherche la signature d'un compte
#  Paramètres:
#   - signName : nom de la signature
#  Retourne un tableau associatif de la signature (id, name, content)
    [Object]searchAccountSignature([String]$signName){
        # Appel méthode getAccountSignature
        $signatures = $this.getAccountSignatures()
        if($signatures -ne $False){
            ForEach($signature in $signatures){
                if($signature.name -eq $signName){
                    return @{"id"=$signature.id;"name"=$signature.name;"type"=$signature.type;"content"=$signature.content}
                }
            }
        }
        return $False
    }

###############################################################
# Méthode createAccountSignature()
#  Ajoute une nouvelle signature au compte
#  Paramètres:
#   - signName : nom de la nouvelle signature
#   - signType : type de la nouvelle signature (text/html)
#   - signContent : contenu de la nouvelle signature
#  Retourne l'Id de la nouvelle signature ou False si échec
    [Object]createAccountSignature([String]$signName, [String]$signType, [String]$signContent){
        # Construction Requête SOAP
        $context = "<urn:account by=`"name`">"+($this.name)+"</urn:account>"
        $request = "<soapenv:Body><urn1:CreateSignatureRequest>"
        $request += "<urn1:signature name=`"$signName`">"        if($signType -eq "text/html"){            $request += "<urn1:content type=`"$signType`"><![CDATA[$signContent]]></urn1:content>"}        if($signType -eq "text/plain"){            $request += "<urn1:content type=`"$signType`">$signContent</urn1:content>"}        $request += "</urn1:signature></urn1:CreateSignatureRequest></soapenv:Body>"
        $response = $this.request($context,$request)

        # Test de la réponse SOAP et renvoie id de la signature ou False si erreur
        if(($response.response) -ne "Error"){
            if($response.Envelope.Body.CreateSignatureResponse.signature.name -eq $signName){
                return $response.Envelope.Body.CreateSignatureResponse.signature.id
            }else{
                return $False
            }
        }else{
            return $False
        }
    }

###############################################################
# Méthode modifyAccountSignature()
#  Modifie une signature du compte
#  Paramètres:
#   - signName : nom de la nouvelle signature
#   - signType : type de la nouvelle signature (text/html)
#   - signContent : contenu de la nouvelle signature
#  Retourne l'Id de la nouvelle signature ou False si échec
    [Object]modifyAccountSignature([String]$signId, [String]$signType, [String]$signContent){
        # Construction Requête SOAP
        $context = "<urn:account by=`"name`">"+($this.name)+"</urn:account>"
        $request = "<soapenv:Body><urn1:ModifySignatureRequest>"
        $request += "<urn1:signature id=`"$signId`">"        $request += "<urn1:content type=`"$signType`"><![CDATA[$signContent]]></urn1:content>"        $request += "</urn1:signature></urn1:ModifySignatureRequest></soapenv:Body>"
        $response = $this.request($context,$request)

        # Test de la réponse SOAP et renvoie id de la signature ou False si erreur
        if(($response.response) -ne "Error"){
            return $True
        }else{
            return $False
        }
    }

###############################################################
# Méthode deleteAccountSignature()
#  Supprime une signature du compte
#  Paramètres:
#   - signName : nom de la signature à supprimer
#  Retourne $True ou $False si échec
    [Object]deleteAccountSignature([String]$signId){
        # Construction Requête SOAP
        $context = "<urn:account by=`"name`">"+($this.name)+"</urn:account>"
        $request = "<soapenv:Body><urn1:DeleteSignatureRequest>"
        $request += "<urn1:signature id=`"$signId`"/>"        $request += "</urn1:DeleteSignatureRequest></soapenv:Body>"
        $response = $this.request($context,$request)

        # Test de la réponse SOAP et renvoie id de la signature ou False si erreur
        if(($response.response) -ne "Error"){
            return $True
        }else{
            return $False
        }
    }

###############################################################
# Méthode getAccountIdentities()
#  Liste les identités du compte
#  Paramètres:
#   - attributs : Liste des attributs des identités à extraire
#  Retourne Tableau d'objet contenant les comptes
    [Object]getAccountIdentities([String[]] $attributs){
        # Construction Requête SOAP
        $context = "<urn:account by=`"name`">"+($this.name)+"</urn:account>"
        $request = "<soapenv:Body><urn1:GetIdentitiesRequest />"
        $request += "</soapenv:Body>"
        $response = $this.request($context,$request)

        # Test de la réponse SOAP et renvoie résultat XML ou False si erreur
        if(($response.response) -ne "Error"){
            $identites = $response.Envelope.Body.GetIdentitiesResponse.identity
            return $this.xmlToObjects($identites,("name","id"),$attributs)
        }else{
            return $False
        }
    }

###############################################################
# Méthode setAccountSignatureDefault()
#  Modifie la signature par défaut du compte
#  Paramètres:
#   - identId : id de l'identité cible
#   - signId : id de la signature
#  Retourne $True ou False si échec
    [Object]setAccountSignatureDefault([String]$identId, [String]$signId){
        # Construction Requête SOAP
        $context = "<urn:account by=`"name`">"+($this.name)+"</urn:account>"
        $request = "<soapenv:Body><urn1:ModifyIdentityRequest>"
        $request += "<urn1:identity id=`"$identId`">"        $request += "<a name=`"zimbraPrefDefaultSignatureId`">$signId</a>"        $request += "</urn1:identity></urn1:ModifyIdentityRequest></soapenv:Body>"
        $response = $this.request($context,$request)

        # Test de la réponse SOAP et renvoie $True ou $False si erreur
        if(($response.response) -ne "Error"){
            return $True
        }else{
            return $False
        }
    }

###############################################################
# Méthode setAccountSignatureReply()
#  Modifie la signature de retour du compte
#  Paramètres:
#   - identId : id de l'identité cible
#   - signId : id de la signature
#  Retourne $True ou False si échec
    [Object]setAccountSignatureReply([String]$identId, [String]$signId){
        # Construction Requête SOAP
        $context = "<urn:account by=`"name`">"+($this.name)+"</urn:account>"
        $request = "<soapenv:Body><urn1:ModifyIdentityRequest>"
        $request += "<urn1:identity id=`"$identId`">"        $request += "<a name=`"zimbraPrefForwardReplySignatureId`">$signId</a>"        $request += "</urn1:identity></urn1:ModifyIdentityRequest></soapenv:Body>"
        $response = $this.request($context,$request)

        # Test de la réponse SOAP et renvoie $True ou $False si erreur
        if(($response.response) -ne "Error"){
            return $True
        }else{
            return $False
        }
    }

###############################################################
# Méthode removeAccountSignatureDefault()
#  Supprime la signature par défaut du compte
#  Paramètres:
#   - identId : id de l'identité cible
#  Retourne $True ou False si échec
    [Object]removeAccountSignatureDefault([String]$identId){
        $signId = "11111111-1111-1111-1111-111111111111"
        # Construction Requête SOAP
        $context = "<urn:account by=`"name`">"+($this.name)+"</urn:account>"
        $request = "<soapenv:Body><urn1:ModifyIdentityRequest>"
        $request += "<urn1:identity id=`"$identId`">"        $request += "<a name=`"zimbraPrefDefaultSignatureId`">$signId</a>"        $request += "</urn1:identity></urn1:ModifyIdentityRequest></soapenv:Body>"
        $response = $this.request($context,$request)

        # Test de la réponse SOAP et renvoie $True ou $False si erreur
        if(($response.response) -ne "Error"){
            return $True
        }else{
            return $False
        }
    }

###############################################################
# Méthode removeAccountSignatureReply()
#  Supprime la signature de retour du compte
#  Paramètres:
#   - identId : id de l'identité cible
#  Retourne $True ou False si échec
    [Object]removeAccountSignatureReply([String]$identId){
        $signId = "11111111-1111-1111-1111-111111111111"
        # Construction Requête SOAP
        $context = "<urn:account by=`"name`">"+($this.name)+"</urn:account>"
        $request = "<soapenv:Body><urn1:ModifyIdentityRequest>"
        $request += "<urn1:identity id=`"$identId`">"        $request += "<a name=`"zimbraPrefForwardReplySignatureId`">$signId</a>"        $request += "</urn1:identity></urn1:ModifyIdentityRequest></soapenv:Body>"
        $response = $this.request($context,$request)

        # Test de la réponse SOAP et renvoie $True ou $False si erreur
        if(($response.response) -ne "Error"){
            return $True
        }else{
            return $False
        }
    }

}