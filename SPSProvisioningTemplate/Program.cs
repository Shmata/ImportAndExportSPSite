using Microsoft.SharePoint.Client;
using Microsoft.IdentityModel.Tokens;


namespace SPSProvisioningTemplate
{
    public class Program
    {
        private const string SOURCE_URL = "https://m365b582028.sharepoint.com/sites/NewSource";
        private const string TARGET_URL = "https://m365b582028.sharepoint.com/sites/Destination";
        static void Main(string[] args)
        {

            var accessToken = TokenHelper.GetAccessToken();
            accessToken.Wait();
            string token = accessToken.Result;
            
            using (var ctx = new ClientContext(SOURCE_URL))
            {
                ctx.ExecutingWebRequest += (sender, args) =>
                {
                    args.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + token;
                };


                ctx.ExecuteQuery();
                if (ctx.Url.IsNullOrEmpty())
                {
                    throw new ArgumentException($"The web Url is not available. Please use context.Load(context.Web, w => w.Url) and context.ExecuteQuery() first.", nameof(ctx));
                }

                ctx.RequestTimeout = Timeout.Infinite;

                SharePointContext services = new SharePointContext(ctx);
                services.SaveSiteAsTemplate(SOURCE_URL);
                services.ApplyTemplate(TARGET_URL);
                services.MoveFiles(TARGET_URL, "SiteAssets");
                //var pallete = services.GetThemePalette(SOURCE_URL, token);
                services.CopyTheme(TARGET_URL, SOURCE_URL, token);

            }


        }

    }
}