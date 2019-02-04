###############################################################
## PowerShell Zimbra Administration
## Version 0.1 - 04/02/2019
##
## Classe ZimbraAdmin permettant des appels SOAP aux API
## d'administration de Zimbra.
##
## Développée et testée sous Zimbra 8.8
## Nécessite PowerShell 6 (pour ignorer les erreurs de certificat)
##
## Auteur : Stéphane GOSNET
##
###############################################################

###############################################################
## URL des WSDL Zimbra Admin
## https://zimbra-serveur:7071/service/wsdl/ZimbraAdminService.wsdl

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
# Méthode testMailExist()
#  Vérifie si une adresse (compte ou alias) existe dans Zimbra
#  Paramètres:
#   - mail : Filtre sur le mail

#  Retourne Tableau d'objet contenant les comptes
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
# Méthode getAccountsByMail()
#  Recherche les comptes à partir du mail
#  Paramètres:
#   - filtre : Filtre sur le mail
#  Retourne Tableau d'objet contenant les comptes
    [Object]getAccountsByMail([String] $filtre,[String[]] $attributs){
        # Construction Requête SOAP
        $attribs = $attributs -join ","
        $request = "<soapenv:Body><urn1:SearchAccountsRequest query=`"(mail=$filtre)`" attrs=`"$attribs`"/></soapenv:Body>"
        $response = $this.request($request)

        # Test de la réponse SOAP et renvoie résultat XML ou False si erreur
        if(($response.response) -ne "Error"){
            [int]$somme = $response.Envelope.Body.SearchAccountsResponse.searchTotal
            $comptes = $response.Envelope.Body.SearchAccountsResponse.account
            return $comptes
        }else{
            return $False
        }
    }


###############################################################
# Méthode getAccountsByCompany()
#  Recherche les comptes appartenant à une société
#  Paramètres:
#   - filtre : Filtre sur la société
#  Retourne Tableau d'objet contenant les comptes
    [Object]getAccountsByCompany([String] $filtre,[String[]] $attributs){
        # Construction Requête SOAP
        $attribs = $attributs -join ","
        $request = "<soapenv:Body><urn1:SearchAccountsRequest query=`"(company=$filtre)`" attrs=`"$attribs`"/></soapenv:Body>"
        $response = $this.request($request)

        # Test de la réponse SOAP et renvoie résultat XML ou False si erreur
        if(($response.response) -ne "Error"){
            [int]$somme = $response.Envelope.Body.SearchAccountsResponse.searchTotal
            $comptes = $response.Envelope.Body.SearchAccountsResponse.account
            return $comptes
        }else{
            return $False
        }
    }


###############################################################
# Méthode getAccountsByTitle()
#  Recherche les comptes à partir de leur fonction
#  Paramètres:
#   - filtre : Filtre sur la fonction
#  Retourne Tableau d'objet contenant les comptes
    [Object]getAccountsByTitle([String] $filtre,[String[]] $attributs){
        # Construction Requête SOAP
        $attribs = $attributs -join ","
        $request = "<soapenv:Body><urn1:SearchAccountsRequest query=`"(title=$filtre)`" attrs=`"$attribs`"/></soapenv:Body>"
        $response = $this.request($request)

        # Test de la réponse SOAP et renvoie résultat XML ou False si erreur
        if(($response.response) -ne "Error"){
            [int]$somme = $response.Envelope.Body.SearchAccountsResponse.searchTotal
            $comptes = $response.Envelope.Body.SearchAccountsResponse.account
            return $comptes
        }else{
            return $False
        }
    }

###############################################################
###############################################################
}


