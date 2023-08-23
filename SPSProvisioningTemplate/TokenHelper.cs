using Microsoft.Identity.Client;
using Microsoft.IdentityModel.Tokens;
using System;
using System.Collections.Generic;
using System.IdentityModel.Tokens.Jwt;
using System.Linq;
using System.Security.Claims;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;

namespace SPSProvisioningTemplate
{
    public static class TokenHelper
    {
        public const string loginUrl = "https://login.microsoftonline.com/";
        public const string tenantId = "c061622d-8f14-4a98-9b67-46d06e27df6c";
        public const string clientId = "b8528544-c8aa-4392-8a38-4990b1406564";

        public static async Task<string> GetAccessToken()
        {

            //var _certificate = X509Certificate2.CreateFromCertFile(@"C:\Host\PnPCore\pk.pfx");
            var _path = @"C:\Host\PnPCore\pk.pfx";
            string certificatePassword = "123";
            var _certificate = new X509Certificate2(System.IO.File.ReadAllBytes(_path), certificatePassword);
            string authority = loginUrl+tenantId;
            var app = ConfidentialClientApplicationBuilder
                                    .Create(clientId)
                                    .WithAuthority(authority, false)
                                    .WithCertificate(_certificate as X509Certificate2)
                                    .Build();



            var token = await app.AcquireTokenForClient(new[] { "https://YOURTenant.sharepoint.com/.default" }).ExecuteAsync();
            return token.AccessToken;

        }


    }
}
