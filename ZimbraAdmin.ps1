###############################################################
## PowerShell Zimbra Administration
## Version 0.1 - 05/06/2019
##
## Classe ZimbraAdmin permettant des appels SOAP aux API
## d'administration de Zimbra.
##
## Développée et testée sous Zimbra 8.8
## PowerShell 5 
## Nécessite PowerShell 6 pour ignorer les erreurs de certificat
##
## Auteur : Stéphane GOSNET
##
###############################################################

###############################################################
## URL des WSDL Zimbra Admin
## https://zimbra-serveur:7071/service/wsdl/ZimbraAdminService.wsdl

###############################################################
## Problèmes à résoudre :
##  - Encodage des accents dans les réponses XML

###############################################################
# Constructeur class
#  Instancie la classe ZimbraAdmin
#  Paramètres :
#   - url : URL du serveur Zimbra (https://zimbra.domaine.com)
#   - $port : Port TCP (7071 par défaut)
#   - ignoreCertificateError : Flag pour ignorer les erreurs de certificat (CA inconnue, date expirée) - Uniquement sous PowerShell 6.1
class ZimbraAdmin
{
hidden [String] $soapHeader = '<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:urn="urn:zimbra" xmlns:urn1="urn:zimbraAdmin"><soapenv:Header><urn:context></urn:context></soapenv:Header>'
hidden [String] $soapFoot = '</soapenv:Envelope>'
hidden [String] $urlSoap = $null
hidden [String] $token = $null

    [String] $url
    [Int] $port = 7071
    [Boolean] $ignoreCertificateError

    ZimbraAdmin([String]$url,[Int]$port,[Switch]$ignoreCertificateError){

        $this.url = $url
        $this.port = $port
        $this.urlSoap = "$url"+":"+"$port/service/admin/soap"
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
    [xml] request([String] $soap){
        
        # Construction de la requete SOAP avec entête et Pied XML
        $request = [xml](($this.soapHeader)+$soap+($this.soapFoot))
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
        $request = "<soapenv:Body><urn1:AuthRequest password=`"$password`"><urn1:account by=`"adminName`">$login</urn1:account></urn1:AuthRequest></soapenv:Body>"
        $response = $this.request($request)

        # Enregistrement du Token et ajout dans l'entête SOAP
        if(($response.response) -ne "Error"){
            $this.token = $response.Envelope.Body.Authresponse.authToken
            $this.soapHeader = '<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:urn="urn:zimbra" xmlns:urn1="urn:zimbraAdmin"><soapenv:Header><urn:context><urn:authToken>'+($this.token)+'</urn:authToken></urn:context></soapenv:Header>'
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
# Méthode delegateAuth()
#  Ouvre une session Admin déléguée Zimbra pour un compte
#  Initie un delegated token Zimbra
#  Paramètres :
#   - name : Compte délégué
#  Retourne Delegated Token ou $False si Erreur
    [String] delegateAuth([String]$name,[Int]$duration){

        # Requête SOAP d'Authentification
        $request = "<soapenv:Body><urn1:DelegateAuthRequest duration=`"$duration`"><urn1:account by=`"name`">$name</urn1:account></urn1:DelegateAuthRequest></soapenv:Body>"
        $response = $this.request($request)

        # Enregistrement du Token et ajout dans l'entête SOAP
        if(($response.response) -ne "Error"){
            return [String]($response.Envelope.Body.DelegateAuthResponse.authToken)
        }else{
            return $False
        }
    }

###############################################################
# Méthode checkHealth()
#  Vérifie l'état de santé de Zimbra
#  Retourne 1 (OK) ou 0 (NOK) selon l'état de santé
    [Boolean]checkHealth(){
        # Construction Requête SOAP
        $request = "<soapenv:Body><urn1:CheckHealthRequest/></soapenv:Body>"
        $response = $this.request($request)

        # Test de la réponse SOAP et renvoie résultat XML ou False si erreur
        if(($response.response) -ne "Error"){
            if($response.Envelope.Body.CheckHealthResponse.healthy -eq "1"){return $true}else{return $false}
        }else{
            return $False
        }
    }

###############################################################
###############################################################
# Méthodes liées aux comptes
###############################################################
###############################################################

###############################################################
# Méthode countAccountsByCos()
#  Compte le nombre de comptes par COS
#  Paramètres:
#   - domaine : domaine de recherche
#  Retourne Tableau d'objet contenant les comptes
    [Object]countAccountsByCos([String] $domain){
        # Construction Requête SOAP
        $request = "<soapenv:Body><urn1:CountAccountRequest><urn1:domain by=`"name`">$domain</urn1:domain></urn1:CountAccountRequest></soapenv:Body>"        $response = $this.request($request)

        # Test de la réponse SOAP et renvoie d'un tableau du nombre de comptes par COS
        if(($response.response) -ne "Error"){
            $coss = $response.Envelope.Body.CountAccountResponse
            ([String]$coss)
            $reponse = @()
            foreach($cos in $coss.cos){
                $reponse += @{($cos.name)=($cos.'#text')}
            }
            return $reponse
        }else{
            return $False
        }
    }

###############################################################
# Méthode testMailExist()
#  Vérifie si une adresse (compte ou alias) existe dans Zimbra
#  Paramètres:
#   - mail : Filtre sur le mail
#  Retourne le nobmre de comptes ou False
    [Boolean]testMailExist([String] $mail){
        # Requête SOAP d'Authentification
        $response = $this.getAccountsByMail($mail,@())
        # Enregistrement du Token et ajout dans l'entête SOAP
        if(($response)){
            return $response.length
        }else{
            return $False
        }
    }

###############################################################
# Méthode getAccountsLocked()
#  Recherche des comptes verouillés
#  Paramètres:
#   - attributs : Liste des attributs à extraire
#  Retourne Tableau d'objet contenant les comptes
    [Object]getAccountsLocked([String[]] $attributs){
        # Construction Requête SOAP
        $attribs = $attributs -join ","
        $request = "<soapenv:Body><urn1:SearchAccountsRequest query=`"(zimbraAccountStatus=locked)`" attrs=`"$attribs`"/></soapenv:Body>"
        $response = $this.request($request)

        # Test de la réponse SOAP et renvoie résultat XML ou False si erreur
        if(($response.response) -ne "Error"){
            $comptes = $response.Envelope.Body.SearchAccountsResponse.account
            return $this.xmlToObjects($comptes,("name","id"),$attributs)
        }else{
            return $False
        }
    }

###############################################################
# Méthode getAccountId()
#  Recherche l'Id d'un compte
#  Paramètres:
#   - name : compte Zimbra
#  Retourne l'Id du compte ou False si échec
    [Object]getAccountId([String] $name){
        # Construction Requête SOAP
        $request = "<soapenv:Body><urn1:SearchAccountsRequest query=`"(mail=$name)`"/></soapenv:Body>"
        $response = $this.request($request)

        # Test de la réponse SOAP et renvoie résultat XML ou False si erreur
        if(($response.response) -ne "Error"){
            return $response.Envelope.Body.SearchAccountsResponse.account.id
        }else{
            return $False
        }
    }

###############################################################
# Méthode getAccountsByMail()
#  Recherche les comptes à partir du mail
#  Paramètres:
#   - filtre : Filtre sur le mail
#   - attributs : Liste des attributs à extraire
#  Retourne Tableau d'objet contenant les comptes
    [Object]getAccountsByMail([String] $filtre,[String[]] $attributs){
        # Construction Requête SOAP
        $attribs = $attributs -join ","
        $request = "<soapenv:Body><urn1:SearchAccountsRequest query=`"(mail=$filtre)`" attrs=`"$attribs`"/></soapenv:Body>"
        $response = $this.request($request)

        # Test de la réponse SOAP et renvoie résultat XML ou False si erreur
        if(($response.response) -ne "Error"){
            $comptes = $response.Envelope.Body.SearchAccountsResponse.account
            return $this.xmlToObjects($comptes,("name","id"),$attributs)
        }else{
            return $False
        }
    }

###############################################################
# Méthode getAccountsByCompany()
#  Recherche les comptes appartenant à une société
#  Paramètres:
#   - filtre : Filtre sur la société
#   - attributs : Liste des attributs à extraire
#  Retourne Tableau d'objet contenant les comptes
    [Object]getAccountsByCompany([String] $filtre,[String[]] $attributs){
        # Construction Requête SOAP
        $attribs = $attributs -join ","
        $request = "<soapenv:Body><urn1:SearchAccountsRequest query=`"(company=$filtre)`" attrs=`"$attribs`"/></soapenv:Body>"
        $response = $this.request($request)

        # Test de la réponse SOAP et renvoie résultat XML ou False si erreur
        if(($response.response) -ne "Error"){
            $comptes = $response.Envelope.Body.SearchAccountsResponse.account
            return $this.xmlToObjects($comptes,("name","id"),$attributs)
        }else{
            return $False
        }
    }

###############################################################
# Méthode getAccountsByTitle()
#  Recherche les comptes à partir de leur fonction
#  Paramètres:
#   - filtre : Filtre sur la fonction
#   - attributs : Liste des attributs à extraire
#  Retourne Tableau d'objet contenant les comptes
    [Object]getAccountsByTitle([String] $filtre,[String[]] $attributs){
        # Construction Requête SOAP
        $attribs = $attributs -join ","
        $request = "<soapenv:Body><urn1:SearchAccountsRequest query=`"(title=$filtre)`" attrs=`"$attribs`"/></soapenv:Body>"
        $response = $this.request($request)

        # Test de la réponse SOAP et renvoie résultat XML ou False si erreur
        if(($response.response) -ne "Error"){
            $comptes = $response.Envelope.Body.SearchAccountsResponse.account
            return $this.xmlToObjects($comptes,("name","id"),$attributs)
        }else{
            return $False
        }
    }

###############################################################
# Méthode getAccountsInGal()
#  Recherche les comptes à partir de leur fonction
#  Paramètres:
#   - name : Filtre de recherche
#   - domaine : domaine de recherche
#  Retourne Tableau d'objet contenant les comptes
    [Object]getAccountsInGal([String] $name,[String] $domain,[String[]]$attributs){
        # Construction Requête SOAP
        $request = "<soapenv:Body><urn1:AutoCompleteGalRequest domain=`"$domain`" name=`"$name`"/></soapenv:Body>"        $response = $this.request($request)

        # Test de la réponse SOAP et renvoie résultat XML ou False si erreur
        if(($response.response) -ne "Error"){
            $comptes = $response.Envelope.Body.AutoCompleteGalResponse.cn
            return $this.xmlToObjects($comptes,("ref","id"),$attributs)
        }else{
            return $False
        }
    }

###############################################################
# Méthode getAccountsMbxSize()
#  Renvoie la taille de la BAL d'un compte
#  Paramètres:
#   - id : Filtre sur l'Id du compte
#   - attributs : Liste des attributs à extraire
#  Retourne Tableau d'objet contenant les comptes
    [Object]getAccountMbxSize([String] $id){
        # Construction Requête SOAP
        $request = "<soapenv:Body><urn1:GetMailboxRequest><urn1:mbox id=`"$id`"/></urn1:GetMailboxRequest></soapenv:Body>"
        $response = $this.request($request)

        # Test de la réponse SOAP et renvoie résultat XML ou False si erreur
        if(($response.response) -ne "Error"){
            return $comptes = $response.Envelope.Body.GetMailboxResponse.mbox.s
        }else{
            return $False
        }
    }

###############################################################
# Méthode addAccount()
#  Creation de compte
#  Paramètres:
#   - name : compte à créer
#   - displayName : Nom affiché compte à créer
#   - sn : Nom du compte à créer
#   - givenName : Prénom du compte à créer
#   - attributs : Tableau associatif des attributs du compte
#  Retourne l'Id du compte ou False si échec
    [Object]addAccount([String]$name, [String]$displayName, [String]$sn, [String]$givenName, $attributs){
        # Construction Requête SOAP
        $request = "<soapenv:Body><urn1:CreateAccountRequest name=`"$name`">"
        $request += "<urn1:a n=`"displayName`">$displayName</urn1:a>"
        $request += "<urn1:a n=`"sn`">$sn</urn1:a>"
        $request += "<urn1:a n=`"givenName`">$givenName</urn1:a>"  
        foreach($key in $attributs.keys){
            $request += "<urn1:a n=`"$key`">"+$attributs[$key]+"</urn1:a>"  
        }
        $request += "</urn1:CreateAccountRequest></soapenv:Body>"

        $response = $this.request($request)
        # Test de la réponse SOAP et renvoie résultat XML ou False si erreur
        if(($response.response) -ne "Error"){
            if($response.Envelope.Body.CreateAccountResponse.account.id){
                return $response.Envelope.Body.CreateAccountResponse.account.id
            }else{
                return $False
            }
        }else{
            return $False
        }
    }

###############################################################
# Méthode renameAccount()
#  Renomme un compte
#  Paramètres:
#   - id : Id du compte
#   - new : Nouveau nom du compte
#  Retourne True ou False si échec
    [Object]renameAccount([String]$id, [String]$new){
        # Construction Requête SOAP
        $request = "<soapenv:Body><urn1:RenameAccountRequest id=`"$id`" newName=`"$new`"/></soapenv:Body>"
        $response = $this.request($request)
        # Test de la réponse SOAP et renvoie résultat XML ou False si erreur
        if(($response.response) -ne "Error"){
            if($response.Envelope.Body.RenameAccountResponse.account.id -eq $id){
                return $true
            }else{
                return $False
            }
        }else{
            return $False
        }
    }

###############################################################
# Méthode modifyAccount()
#  Modifiie les attributs d'un compte
#  Paramètres:
#   - id : Id du compte
#   - attributs : tableau associatif des attributs à modifier
#  Retourne True ou False si échec
    [Object]modifyAccount([String]$id, $attributs){
        # Construction Requête SOAP
        $request = "<soapenv:Body><urn1:ModifyAccountRequest id=`"$id`">"
        Foreach($key in ($attributs.keys)){
            $request += "<urn1:a n=`"$key`">"+$attributs[$key]+"</urn1:a>"
        }        
        $request += "</urn1:ModifyAccountRequest></soapenv:Body>"
        $response = $this.request($request)
        # Test de la réponse SOAP et renvoie résultat XML ou False si erreur
        if(($response.response) -ne "Error"){
            if($response.Envelope.Body.ModifyAccountResponse.account.id -eq $id){
                return $true
            }else{
                return $False
            }
        }else{
            return $False
        }
    }

###############################################################
# Méthode deleteAccount()
#  Supprime un compte
#  Paramètres:
#   - id : Id du compte
#  Retourne True ou False si échec
    [Object]deleteAccount([String]$id){
        # Construction Requête SOAP
        $request = "<soapenv:Body><urn1:DeleteAccountRequest id=`"$id`"/></soapenv:Body>"
        $response = $this.request($request)
        # Test de la réponse SOAP et renvoie résultat XML ou False si erreur
        if(($response.response) -ne "Error"){
            if($response.Envelope.Body.DeleteAccountResponse){
                return $true
            }else{
                return $False
            }
        }else{
            return $False
        }
    }

###############################################################
###############################################################
# Méthodes liées aux alias de comptes
###############################################################
###############################################################

###############################################################
# Méthode getAccountAlias()
#  Revnoie les alias d'un compte
#  Paramètres:
#   - id : Id du compte
#  Retourne un tableau des alias ou False si échec
    [Object]getAccountAlias([String]$id){
        # Construction Requête SOAP
        $request = "<soapenv:Body><urn1:GetAccountRequest attrs=`"zimbraMailAlias`">"
        $request += "<urn1:account by=`"id`">$id</urn1:account>"
        $request += "</urn1:GetAccountRequest></soapenv:Body>"
        $response = $this.request($request)

        # Test de la réponse SOAP et renvoie résultat XML ou False si erreur
        if(($response.response) -ne "Error"){
            if($response.Envelope.Body.getAccountResponse.account.id -eq $id){
                $alias = @()
                ForEach($a in $response.Envelope.Body.getAccountResponse.account.a){
                    $alias += $a.'#text'
                }
                return $alias
            }else{
                return $False
            }
        }else{
            return $False
        }
    }

###############################################################
# Méthode addAccountAlias()
#  Ajoute un alais à un compte
#  Paramètres:
#   - id : Id du compte
#   - alais : Alias du compte à créer
#  Retourne True ou False si échec
    [Object]addAccountAlias([String]$id, [String]$alias){
        # Construction Requête SOAP
        $request = "<soapenv:Body><urn1:AddAccountAliasRequest id=`"$id`" alias=`"$alias`"/></soapenv:Body>"
        $response = $this.request($request)
        # Test de la réponse SOAP et renvoie résultat XML ou False si erreur
        if(($response.response) -ne "Error"){
            if($response.Envelope.Body.AddAccountAliasResponse.xmlns){
                return $true
            }else{
                return $False
            }
        }else{
            return $False
        }
    }

###############################################################
# Méthode removeAccountAlias()
#  Supprime un alais à un compte
#  Paramètres:
#   - id : Id du compte
#   - alais : Alias du compte à supprimer
#  Retourne True ou False si échec
    [Object]removeAccountAlias([String]$id, [String]$alias){
        # Construction Requête SOAP
        $request = "<soapenv:Body><urn1:RemoveAccountAliasRequest id=`"$id`" alias=`"$alias`"/></soapenv:Body>"
        $response = $this.request($request)
        # Test de la réponse SOAP et renvoie résultat XML ou False si erreur
        if(($response.response) -ne "Error"){
            if($response.Envelope.Body.RemoveAccountAliasResponse.xmlns){
                return $true
            }else{
                return $False
            }
        }else{
            return $False
        }
    }


###############################################################
###############################################################
}


