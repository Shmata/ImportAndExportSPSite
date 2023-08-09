# ImportAndExportSPSite
You can easily create an instance (a template) from your SharePoint online site including lists,docs,settings,users,content types and export it to another SharePoint online site. 

## First create a self signed certificate
in powershell run below command 

**$cert=New-SelfSignedCertificate -Subject "CN=Sapiens.at.SharePoint.Shahab" -CertStoreLocation "Cert:\CurrentUser\My"  -KeyExportPolicy Exportable -KeySpec Signature -NotAfter (Get-Date).AddMonths(24)**

## Second export your recently created certificate in addition to private key 
in run -> mmc -> Certificates - Current User -> Personal -> Certificate 
in certificate right click select All tasks then export 

## Third create an app registeration in Azure portal 
In the Certificates and Sectrets -> select certificates -> upload your self signed certificate. 

## Fourth get client Id, Tenant Id
Go to overview section and get the client id and tenant id and replace it in console app. 
