using Microsoft.Identity.Client;
using Microsoft.SharePoint.Client;
using PnP.Framework.Provisioning.Connectors;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;
using PnP.Framework.Provisioning.Providers.Xml;
using System;
using System.IO;
using System.Net;
using System.Security;
using System.Security.Cryptography.X509Certificates;
using System.Threading;


//Console.WriteLine("Hello, World!");

namespace SPSProvisioningTemplate
{
    class Program
    {
        
        static void Main(string[] args)
        {
            //ConsoleColor defaultForeground = Console.ForegroundColor;
            //Console.ForegroundColor = ConsoleColor.Green;


            var accessToken = GetAccessToken();
            accessToken.Wait();
            string token = accessToken.Result;
            string sourceSite = "https://tenant.sharepoint.com/sites/subsite";
            
            GetProvisioningTemplate(sourceSite, token);

        }

        private static async Task<string> GetAccessToken()
        {

            //var _certificate = X509Certificate2.CreateFromCertFile(@"C:\Host\PnPCore\pk.pfx");
            var _path = @"C:\Host\PnPCore\pk.pfx";
            string certificatePassword = "123";
            var _certificate = new X509Certificate2(System.IO.File.ReadAllBytes(_path), certificatePassword);
            string authority = $"https://login.microsoftonline.com/_TenantID";
            var app = ConfidentialClientApplicationBuilder
                                    .Create("b8528544-c8aa-4392-8a38-4990b1406564")
                                    .WithAuthority(authority, false)
                                    .WithCertificate(_certificate as X509Certificate2)
                                    .Build();



            var token = await app.AcquireTokenForClient(new[] { "https://tenant.sharepoint.com/.default" }).ExecuteAsync();
            return token.AccessToken;
            
        }

        private static ProvisioningTemplate GetProvisioningTemplate( string webUrl, string accessToken)
        {
            using (var ctx = new ClientContext(webUrl))
            {
                ctx.ExecutingWebRequest += (sender, args) =>
                {
                    args.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + accessToken;
                };


                //ctx.Credentials = new SharePointOnlineCredentials(userName, pwd);
                ctx.RequestTimeout = Timeout.Infinite;

                // Just to output the site details
                Web web = ctx.Web;
                ctx.Load(web, w => w.Title);
                ctx.ExecuteQueryRetry();

                //Console.ForegroundColor = ConsoleColor.White;
                //Console.WriteLine("Your site title is:" + ctx.Web.Title);
                //Console.ForegroundColor = defaultForeground;

                ProvisioningTemplateCreationInformation ptci
                        = new ProvisioningTemplateCreationInformation(ctx.Web);

                ptci.FileConnector = new FileSystemConnector(@"c:\Host\PnPCore", "");
                ptci.PersistComposedLookFiles = true;
                ptci.ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
                {
                    Console.WriteLine("{0:00}/{1:00} - {2}", progress, total, message);
                };

                // Execute actual extraction of the template
                ProvisioningTemplate template = ctx.Web.GetProvisioningTemplate(ptci);

                // We can serialize this template to save and reuse it
                XMLTemplateProvider provider =
                        new XMLFileSystemTemplateProvider(@"c:\Host\PnPCore", "");
                provider.SaveAs(template, "PnPProvisioningDemo.xml");
                
                ApplyProvisioningTemplate(template, accessToken);
                return template;
            }
        }


        private static void ApplyProvisioningTemplate( ProvisioningTemplate template, string accessToken)
        {
            string targetSite = "https://tenant.sharepoint.com/sites/distinationSite";
            using (var ctx = new ClientContext(targetSite))
            {
                ctx.ExecutingWebRequest += (sender, args) =>
                {
                    args.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + accessToken;
                };
                ctx.RequestTimeout = Timeout.Infinite;

                Web web = ctx.Web;

                ProvisioningTemplateApplyingInformation ptai = new ProvisioningTemplateApplyingInformation();
                ptai.ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
                {
                    Console.WriteLine("{0:00}/{1:00} - {2}", progress, total, message);
                };

                // Associate file connector for assets
                FileSystemConnector connector = new FileSystemConnector(@"c:\Host\PnPCore\PnPProvisioningDemo", "");
                template.Connector = connector;

                // Because the template is actual object, we can modify this using code as needed
                template.Lists.Add(new ListInstance()
                {
                    Title = "Test List",
                    Url = "lists/Testlist",
                    TemplateType = (Int32)ListTemplateType.Contacts,
                    EnableAttachments = true
                });

                web.ApplyProvisioningTemplate(template, ptai);
            }
        }

    }
}
