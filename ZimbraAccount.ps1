###############################################################
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

    [String] $url
    [Int] $port = 8443
    [Boolean] $ignoreCertificateError

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
                    if($ligne.n -eq $attribut){
                        $attrCnt[$attribut] += 1
                    }
                }
            }
            # Traitement des attributs
            $attrMulti=@{}
            foreach($ligne in $result.a){
                foreach($attribut in $attributs){
                    if($ligne.n -eq $attribut){
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
            $this.token = $response.Envelope.Body.Authresponse.authToken
            $this.soapHeader1 = '<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:urn="urn:zimbra" xmlns:urn1="urn:zimbraAccount"><soapenv:Header><urn:context><urn:authToken>'+($this.token)+'</urn:authToken>'
            $this.soapHeader2 = '</urn:context></soapenv:Header>'
            return $True
        }else{
            return $False
        }
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
#  Paramètres:
#   - name : nom du compte
#  Retourne un tableau associatif des signatures (id, name, content)
    [Object]getAccountSignatures([String]$compte){
        # Construction Requête SOAP
        $context = "<urn:account by=`"name`">$compte</urn:account>"
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
#   - name : nom du compte
#   - signName : nom de la signature
#  Retourne un tableau associatif de la signature (id, name, content)
    [Object]searchAccountSignature([String]$compte,[String]$signName){
        # Appel méthode getAccountSignature
        $signatures = $this.getAccountSignatures($compte)
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
#   - name : nom du compte
#   - signName : nom de la nouvelle signature
#   - signType : type de la nouvelle signature (text/html)
#   - signContent : contenu de la nouvelle signature
#  Retourne l'Id de la nouvelle signature ou False si échec
    [Object]createAccountSignature([String]$name, [String]$signName, [String]$signType, [String]$signContent){
        # Construction Requête SOAP
        $context = "<urn:account by=`"name`">$name</urn:account>"
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
#   - name : nom du compte
#   - signName : nom de la nouvelle signature
#   - signType : type de la nouvelle signature (text/html)
#   - signContent : contenu de la nouvelle signature
#  Retourne l'Id de la nouvelle signature ou False si échec
    [Object]modifyAccountSignature([String]$name, [String]$signId, [String]$signType, [String]$signContent){
        # Construction Requête SOAP
        $context = "<urn:account by=`"name`">$name</urn:account>"
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
#   - name : nom du compte
#   - signName : nom de la signature à supprimer
#  Retourne $True ou $False si échec
    [Object]deleteAccountSignature([String]$name, [String]$signId){
        # Construction Requête SOAP
        $context = "<urn:account by=`"name`">$name</urn:account>"
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

}