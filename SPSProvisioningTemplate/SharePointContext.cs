using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;
using PnP.Framework.Provisioning.Providers;
using PnP.Framework.Provisioning.Providers.Xml;
using PnP.Framework.Utilities;
using Sapiens.at.SharePoint.QueryBuilder;
using System.Net;
using System.Text.Json.Nodes;
using System.Xml;
using System.Xml.Linq;
using FieldLookupValue = Microsoft.SharePoint.Client.FieldLookupValue;
using FieldType = Microsoft.SharePoint.Client.FieldType;
using FieldUrlValue = Microsoft.SharePoint.Client.FieldUrlValue;
using FieldUserValue = Microsoft.SharePoint.Client.FieldUserValue;
using File = Microsoft.SharePoint.Client.File;

namespace SPSProvisioningTemplate
{
    public class SharePointContext
    {
        private string accessToken;
        private string webUrl = "https://YOURTenant.sharepoint.com/sites/NewSource";
        private const string TEMPLATE_LIBRARY = "SiteAssets";
        private const string TEMPLATE_FILENAME = "PnPProvisioningDemo.xml";
        private const string REPLACE_SLASH_WITH = "__";
        private readonly static Microsoft.SharePoint.Client.FieldType[] _unsupportedFieldTypes =
        {
            FieldType.Attachments,
            FieldType.Computed
        };
        public TemplateProviderBase TemplateProvider { get; }
        public ClientContext Context { get; }
        public SharePointContext(ClientContext contextOfTheSourceSite)
        {
            Context = contextOfTheSourceSite;
            TemplateProvider = new XMLSharePointTemplateProvider(Context, Context.Url, TEMPLATE_LIBRARY);
        }
        // This method is respoinsible to create a template from a source site, the template is in XML format and will be store in the SiteAssets library of the source site. 
        public void SaveSiteAsTemplate(string sourceURL)
        {
            if (string.IsNullOrEmpty(sourceURL))
            {
                throw new ArgumentException($"'{nameof(sourceURL)}' cannot be null or empty.", nameof(sourceURL));
            }

            RenameTemplateFileIfAlreadyExists();

            List<string> listsToExtract;
            
            LoadSharePointProperties(out listsToExtract);

            Context.Load(Context.Web, sw => sw.ServerRelativeUrl, sw => sw.Url, sw => sw.Lists.Include(l => l.Title, l => l.RootFolder.ServerRelativeUrl, l => l.Fields.Include(f => f.InternalName, f => f.Title, f => f.SchemaXml)));
            Context.ExecuteQuery();

            // get the template from the subweb
            ProvisioningTemplateCreationInformation ptci = new ProvisioningTemplateCreationInformation(Context.Web);
            ptci.ProgressDelegate = LogProcessToConsole;
            ptci.MessagesDelegate = LogMessageToConsole;
            ptci.IncludeAllClientSidePages = true;
            ptci.IncludeHiddenLists = true;
            //ptci.PersistBrandingFiles = true;
            ptci.ListsToExtract = listsToExtract;
            ptci.IncludeSiteGroups = true;
            var template = Context.Web.GetProvisioningTemplate(ptci);

            // remove all custom actions
            ClearCustomActions(template);

            // fix calculated columns from all lists
            FixColumns(template, Context.Web.Lists);

            // fix client side pages if they are in a subfolder
            FixClientSidePages(template);

            // add records
            AddListDataToTemplate(Context.Web, template);

            // save the template to the SharePoint library
            TemplateProvider.SaveAs(template, TEMPLATE_FILENAME);
        }
        // As already mentioned the template file would be store in the site assets of source site, if there was another template from another sites, those existing template would be renamed
        private void RenameTemplateFileIfAlreadyExists()
        {
            Context.Load(Context.Web, w => w.ServerRelativeUrl);
            Context.ExecuteQuery();

            string webServerRelativeUrl = Context.Web.ServerRelativeUrl;
            string serverRelativeUrl = $"{EnsureTrailingSlash(webServerRelativeUrl)}{TEMPLATE_LIBRARY}/{TEMPLATE_FILENAME}";

            //string serverRelativeUrl = $"{EnsureTrailingSlash(Context.Web.ServerRelativeUrl)}{TEMPLATE_LIBRARY}/{TEMPLATE_FILENAME}";
            string timeStamp = DateTime.Now.ToString("yyyy-MM-dd_hhmmss");
            string serverRelativeUrlIfExists = $"{EnsureTrailingSlash(Context.Web.ServerRelativeUrl)}{TEMPLATE_LIBRARY}/{timeStamp}_{TEMPLATE_FILENAME}";

            // get the template file
            var templateFile = Context.Web.GetFileByServerRelativeUrl(serverRelativeUrl);
            Context.Load(templateFile, t => t.Exists);
            Context.ExecuteQuery();

            // move if already exists
            if (templateFile.Exists)
            {
                templateFile.MoveTo(serverRelativeUrlIfExists, Microsoft.SharePoint.Client.MoveOperations.None);
                Context.ExecuteQuery();
            }

        }
        private void LoadSharePointProperties(out List<string> listsToExtract)
        {
            IEnumerable<File> siteAssets;
            // get all lists created by the EVM or ETM app
            var excludeLists = new string[] { "Style Library", "FormServerTemplates", "Shared%20Documents", "Shared Documents", "masterpage", "solutions", "wp",
                "Converted Forms", "Documents", "theme", "SiteAssets", "Composed Looks", "Composed Looks", "appdata", "appfiles", "lt", "design", "IWConvertedForms",
                 "wte", "SitePages", "catalogs/" };
            //var includeLists = new string[] { "Lists" };
            Context.Load(Context.Web, w => w.Lists.Include(l => l.RootFolder.ServerRelativeUrl, l => l.Title));
            Context.Load(Context.Web, c => c.ServerRelativeUrl);
            Context.ExecuteQuery();
            //getlist
            //listsToExtract = Context.Web.Lists.Where(l => includeLists.Any(il => l.RootFolder.ServerRelativeUrl.Contains(il))).Select(l => l.Title).ToList();
            listsToExtract = Context.Web.Lists.Where(l => !excludeLists.Any(il => l.RootFolder.ServerRelativeUrl.Contains(il))).Select(l => l.Title).ToList();
            if (listsToExtract.Count == 0)
            {
                throw new Exception($"There is no lists in this site");
            }
            var siteAssetsList = Context.Web.GetList(UrlUtility.Combine(Context.Web.ServerRelativeUrl, "SiteAssets"));
            var siteAssetItems = siteAssetsList.GetItems(new CamlQuery().Scope(Microsoft.SharePoint.Client.ViewScope.Recursive));
            Context.Load(siteAssetItems, sac => sac.Include(sa => sa.File.ServerRelativeUrl, sa => sa.File.Name));
            Context.ExecuteQuery();
            siteAssets = siteAssetItems.Select(i => i.File).ToArray();
        }

        // Custom actions should be removed from the template so this method will do that
        private void ClearCustomActions(ProvisioningTemplate template)
        {
            template.CustomActions.SiteCustomActions.Clear();
            template.CustomActions.WebCustomActions.Clear();
        }
        // In order to copy and restore pages we need to replace / character with __
        private static void FixClientSidePages(ProvisioningTemplate template)
        {
            // fix client side pages (if they are in a subfolder)
            template.ClientSidePages.ToArray().ToList().ForEach((p) =>
            {
                if (p.PageName.Contains("/"))
                {
                    p.PageName = p.PageName.Replace("/", REPLACE_SLASH_WITH);
                }
            });
        }
        // If we make a decision to copy lists with their items this mehtod is responsible to provide items in our lists
        private void AddListDataToTemplate(Web subWeb, ProvisioningTemplate template)
        {
            // get list data
            foreach (var l in template.Lists.ToArray())
            {
                l.DataRows.UpdateBehavior = UpdateBehavior.Overwrite;
                AddListDataPerList(subWeb, l);
            }
        }
        private void AddListDataPerList(Web subWeb, ListInstance l)
        {
            var list = subWeb.Lists.GetByTitle(l.Title);
            var items = list.GetItems(new CamlQuery());
            Context.Load(items);
            Context.Load(list.Fields, fc => fc.Include(f => f.InternalName, f => f.ReadOnlyField, f => f.FieldTypeKind));
            Context.ExecuteQuery();

            var fieldsToExport = list.Fields
                                .Where(f => !f.ReadOnlyField && !_unsupportedFieldTypes.Contains(f.FieldTypeKind));
            foreach (var listItem in items)
            {
                DataRow dataRow = new DataRow();
                foreach (var field in fieldsToExport)
                {
                    var fldKey = (from f in listItem.FieldValues.Keys where f == field.InternalName select f).FirstOrDefault();
                    if (!string.IsNullOrEmpty(fldKey))
                    {
                        var fieldValue = GetFieldValueAsText(subWeb, listItem, field);
                        dataRow.Values.Add(field.InternalName, fieldValue);
                    }
                }
                l.DataRows.Add(dataRow);
            }
        }

        private string GetFieldValueAsText(Web web, ListItem listItem, Microsoft.SharePoint.Client.Field field)
        {
            if (field.InternalName == "sapiensatUpdateID")
            {
                return "1"; // make sure we don't run any receivers.
            }

            var rawValue = listItem[field.InternalName];
            if (rawValue == null) return null;

            // Since the TaxonomyField is not in the FieldTypeKind enumeration below, a specific check is done here for this type
            if (field is TaxonomyField)
            {
                if (rawValue is TaxonomyFieldValueCollection)
                {
                    List<string> termIds = new List<string>();
                    foreach (var taxonomyValue in (TaxonomyFieldValueCollection)rawValue)
                    {
                        termIds.Add($"{taxonomyValue.TermGuid}");
                    }
                    return String.Join(";", termIds);
                }
                else if (rawValue is TaxonomyFieldValue)
                {
                    return $"{((TaxonomyFieldValue)rawValue).TermGuid}";
                }
            }

            // Specific operations based on the type of field at hand
            switch (field.FieldTypeKind)
            {
                case FieldType.Geolocation:
                    var geoValue = (FieldGeolocationValue)rawValue;
                    return $"{geoValue.Altitude},{geoValue.Latitude},{geoValue.Longitude},{geoValue.Measure}";
                case FieldType.URL:
                    var urlValue = (FieldUrlValue)rawValue;
                    return $"{urlValue.Url},{urlValue.Description}";
                case FieldType.Lookup:
                    var strVal = rawValue as string;
                    if (strVal != null)
                    {
                        return strVal;
                    }
                    var singleLookupValue = rawValue as FieldLookupValue;
                    if (singleLookupValue != null)
                    {
                        return singleLookupValue.LookupId.ToString();
                    }
                    var multipleLookupValue = rawValue as FieldLookupValue[];
                    if (multipleLookupValue != null)
                    {
                        return string.Join(",", multipleLookupValue.Select(lv => lv.LookupId));
                    }
                    throw new Exception("Invalid data in field");
                case FieldType.User:
                    var singleUserValue = rawValue as FieldUserValue;
                    if (singleUserValue != null)
                    {
                        return GetLoginName(web, singleUserValue.LookupId);
                    }
                    var multipleUserValue = rawValue as FieldUserValue[];
                    if (multipleUserValue != null)
                    {
                        return string.Join(",", multipleUserValue.Select(lv => GetLoginName(web, lv.LookupId)));
                    }
                    throw new Exception("Invalid data in field");
                case FieldType.MultiChoice:
                    var multipleChoiceValue = rawValue as string[];
                    if (multipleChoiceValue != null)
                    {
                        return string.Join(";#", multipleChoiceValue);
                    }
                    return Convert.ToString(rawValue);
                default:
                    return Convert.ToString(rawValue);
            }
        }
        // This method is the beginning point of restore template to the target site. 
        public void ApplyTemplate(string targetWebUrl)
        {
            var template = TemplateProvider.GetTemplate(TEMPLATE_FILENAME);

            if (template == null)
            {
                throw new Exception($"The template '{TEMPLATE_FILENAME}' does not exist in the library '{TEMPLATE_LIBRARY}' in '{webUrl}'.");
            }

            using (var contextNewWeb = Context.Clone(targetWebUrl))
            {
                var newWeb = contextNewWeb.Web;

                // Apply the template
                ProvisioningTemplateApplyingInformation ptai = new ProvisioningTemplateApplyingInformation();
                ptai.ProgressDelegate = LogProcessToConsole;
                ptai.MessagesDelegate = LogMessageToConsole;
                ptai.ClearNavigation = true;
                ptai.ProvisionContentTypesToSubWebs = true;
                newWeb.ApplyProvisioningTemplate(template, ptai);

                // Fix site pages after provisioning
                FixSitePagesAfterProvisioning(contextNewWeb, newWeb);
            }
        }
        // This is the last method will be executed after restore, in order to run move files and replace _ with / and other fixing thingy
        private static void FixSitePagesAfterProvisioning(ClientContext contextNewWeb, Web web)
        {
            var sitepages = web.GetSitePagesLibrary();
            var allPages = sitepages.GetItems(new CamlQuery());
            contextNewWeb.Load(web, w => w.ServerRelativeUrl);
            contextNewWeb.Load(allPages, p => p.Include(f => f.File.ServerRelativeUrl));
            contextNewWeb.ExecuteQuery();
            var keys = new string[] { "/" };
            foreach (var p in allPages)
            {
                var key = keys.FirstOrDefault(k => p.File.ServerRelativeUrl.ToLower().Contains(k.ToLower()));
                if (!string.IsNullOrEmpty(key))
                {
                    string newUrl = p.File.ServerRelativeUrl;
                    var baseUrl = newUrl.Substring(0, newUrl.ToLower().IndexOf(key.ToLower()));
                    var libraryRelativeUrl = newUrl.Substring(newUrl.ToLower().IndexOf(key.ToLower()));
                    newUrl = UrlUtility.Combine(baseUrl, libraryRelativeUrl.Replace(REPLACE_SLASH_WITH, "/"));
                    EnsureFolderPath(web, newUrl);
                    p.File.MoveTo(newUrl, Microsoft.SharePoint.Client.MoveOperations.None);
                    contextNewWeb.ExecuteQuery();
                }
            }
        }
        // This method is in charge of copy files in to target site physically
        public void MoveFiles(string targetURL, string documentLibrary)
        {
            var template = new ProvisioningTemplate();
            Context.Load(Context.Web, u => u.Url, u => u.ServerRelativeUrl);
            Context.ExecuteQuery();
            template.Connector = new PnP.Framework.Provisioning.Connectors.SharePointConnector(Context, Context.Url, documentLibrary);
            var docLib = Context.Web.GetList(UrlUtility.Combine(Context.Web.ServerRelativeUrl, documentLibrary));
            var files = docLib.GetItems(new CamlQuery().Scope(Microsoft.SharePoint.Client.ViewScope.Recursive));
            Context.Load(files, fc => fc.Include(f => f.File.ServerRelativeUrl));
            Context.ExecuteQuery();
            foreach (var f in files.Select(f => f.File).Where(f => f.ServerObjectIsNull != true))
            {
                var siteRelativeUrl = f.ServerRelativeUrl.Substring(Context.Web.ServerRelativeUrl.Length);
                if (siteRelativeUrl.StartsWith("/")) siteRelativeUrl = siteRelativeUrl.Substring(1);
                // remove doc lib path
                var src = siteRelativeUrl.Substring(siteRelativeUrl.IndexOf("/") + 1);
                // only file name
                var folder = siteRelativeUrl.Substring(0, siteRelativeUrl.LastIndexOf("/"));
                template.Files.Add(new PnP.Framework.Provisioning.Model.File()
                {
                    Src = src,
                    Folder = folder,
                    Overwrite = true,
                    Level = PnP.Framework.Provisioning.Model.FileLevel.Published
                });
            }

            using (var newWebContext = Context.Clone(targetURL))
            {
                ProvisioningTemplateApplyingInformation ptai = new ProvisioningTemplateApplyingInformation();
                ptai.ProgressDelegate = LogProcessToConsole;
                ptai.MessagesDelegate = LogMessageToConsole;
                newWebContext.Web.ApplyProvisioningTemplate(template, ptai);
            }
        }

        // This method will provide theme colors url 
        public string GetThemePaletteUrl(string sourceURL, string token)
        {
            var url = $"{sourceURL}/_api/web/ThemedCssFolderUrl";
            var request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "POST";
            request.ContentType = "application/json;charset=utf-8";
            request.Accept = "application/json;odata.metadata=minimal";
            request.Headers.Add("Authorization", "Bearer " + token);
            request.Headers.Add("ODATA-VERSION", "4.0");
            var response = (HttpWebResponse)request.GetResponse();
            var sr = new StreamReader(response.GetResponseStream());
            string responseText = sr.ReadToEnd();
            //var jsonObject = JsonConvert.DeserializeObject(responseText);
            var jsonObject = JsonConvert.DeserializeObject<JToken>(responseText);
            var themeUrl = jsonObject["value"].ToString();
            Context.Load(Context.Web, w => w.ServerRelativeUrl);
            Context.ExecuteQuery();
            var baseUrl = RemoveSiteFromUrl(Context.Url, Context.Web.ServerRelativeUrl);
            var paletteUrl = baseUrl + themeUrl + "/theme.spcolor";
            return paletteUrl;
        }

        // In order to generate a correct URL to access to the theme.spcolor we need to manipulate the URL
        public string RemoveSiteFromUrl(string fullUrl, string siteUrl)
        {
            if (fullUrl.EndsWith(siteUrl, StringComparison.OrdinalIgnoreCase))
            {
                int index = fullUrl.LastIndexOf(siteUrl, StringComparison.OrdinalIgnoreCase);
                if (index >= 0)
                {
                    return fullUrl.Substring(0, index);
                }
            }

            return fullUrl;
        }
        // theme.spcolor is in XML format in an XML file, we need to provide valid format in json so in this method we will remove extra XML's tags then generate a valid json object
        static string TransformColorPalette(JObject jsonObject)
        {
            var colorPaletteElement = jsonObject["s:colorPalette"];
            if (colorPaletteElement == null)
            {
                throw new InvalidOperationException("s:colorPalette element not found.");
            }
            JObject transformedPalette = new JObject();
            foreach (var colorElement in colorPaletteElement["s:color"])
            {
                string name = colorElement["@name"].ToString();
                string value = '#' + colorElement["@value"].ToString();
                transformedPalette[name] = value;
            }

            JObject result = new JObject(new JProperty("palette", transformedPalette));
            return result.ToString(Newtonsoft.Json.Formatting.Indented);
        }
        // This method will get theme.spcolor xml file and return a valid theme palette in a json format. 
        // We will use a few method to return a correct result.
        public object GetThemePalette(string sourceURL, string token)
        {
            var paletteUrl = GetThemePaletteUrl(sourceURL, token);
            var file = Context.Web.GetFileAsString(new Uri(paletteUrl).LocalPath);
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(file);
            string json = JsonConvert.SerializeXmlNode(doc);
            JObject jsonObject = JObject.Parse(json);
            var validJsonPalette = TransformColorPalette(jsonObject);
            return validJsonPalette;

        }
        // We will get theme.spcolor palette from source site and apply that to the target site. 
        public async void CopyTheme(string targetURL, string sourceURL, string token)
        {

            object palette = GetThemePalette(sourceURL, token);
            var template = new ProvisioningTemplate();
            Context.Load(Context.Web, u => u.ThemeInfo);
            Context.ExecuteQuery();
            var themeEntity = Context.Web.GetCurrentComposedLook();

            using (var newWebContext = Context.Clone(targetURL))
            {
                newWebContext.Load(newWebContext.Web, w => w.ServerRelativeUrl);
                newWebContext.ExecuteQuery();

                var url = $"{targetURL}/_api/thememanager/ApplyTheme";
                var request = (HttpWebRequest)WebRequest.Create(url);
                request.Method = "POST";
                request.ContentType = "application/json;charset=utf-8";
                request.Accept = "application/json;odata.metadata=minimal";
                request.Headers.Add("Authorization", "Bearer " + token);
                request.Headers.Add("ODATA-VERSION", "4.0");

                using (var streamWriter = new StreamWriter(request.GetRequestStream()))
                {
                    //string json = JsonConvert.SerializeObject(new { name = "Sounders Rave Green", themeJson = pal });
                    string json = JsonConvert.SerializeObject(new { name = "Sounders Rave Green", themeJson = palette });
                    streamWriter.Write(json);
                }

                var response = (HttpWebResponse)request.GetResponse();
                using (var sr = new StreamReader(response.GetResponseStream()))
                {
                    string responseText = sr.ReadToEnd();
                    var jsonObject = JsonConvert.DeserializeObject(responseText);
                }

            }
        }

        // 
        private Dictionary<Guid, Dictionary<int, string>> _webUserCache = new Dictionary<Guid, Dictionary<int, string>>();
        private JsonObject jsonObject;
        private string GetLoginName(Web web, int userId)
        {
            if (!_webUserCache.ContainsKey(web.Id)) _webUserCache.Add(web.Id, new Dictionary<int, string>());
            if (!_webUserCache[web.Id].ContainsKey(userId))
            {
                var user = web.GetUserById(userId);
                web.Context.Load(user, u => u.LoginName);
                web.Context.ExecuteQueryRetry();
                _webUserCache[web.Id].Add(userId, user.LoginName);
            }
            return _webUserCache[web.Id][userId];
        }
        // We are not able to restore columns in a target site by default so we have to fix some possible issues using replace some keys and values, we will do that using below method. 
        private void FixColumns(ProvisioningTemplate template, ListCollection lists)
        {
            Context.Load(lists);
            Context.ExecuteQuery();
            // check all lists and fix the formulas for calculated columns
            foreach (var l in template.Lists.ToArray())
            {
                var spList = lists.FirstOrDefault(spl => spl.Title == l.Title);
                FixColumnsPerList(l, spList);
            }
        }
        // Fix calculated columns in addition to remove <Validation> from XML template
        private void FixColumnsPerList(ListInstance l, List spList)
        {
            Dictionary<string, string> replacements = GetReplacements(l, spList);
            // fix calculated columns and columns with a column validation
            l.Fields.Where(f => f.SchemaXml.Contains(" Type=\"Calculated\"") || f.SchemaXml.Contains("<Validation ")).ToArray().ToList().ForEach((f) =>
            {
                FixCalculatedColumns(f, replacements);
            });

            Context.Load(spList.RootFolder, rf => rf.ServerRelativeUrl);
            Context.ExecuteQuery();
            // If EVM is already installed we have to fix some possilbe issues, we will do that using FixColumnsEventList
            if (spList.RootFolder.ServerRelativeUrl.ToLower().Contains("Lists/sapiensEvents".ToLower()))
            {
                FixColumnsInEventList(l, spList);
            }
        }
        private void FixColumnsInEventList(ListInstance l, List spList)
        {
            var removeFieldRefAndAddDirectly = new string[] { "EventDate", "EndDate" };
            var fieldRefsToRemove = new List<FieldRef>();

            foreach (var fr in l.FieldRefs)
            {
                if (removeFieldRefAndAddDirectly.Contains(fr.Name))
                {
                    fieldRefsToRemove.Add(fr);
                }
            }

            foreach (var frToRemove in fieldRefsToRemove)
            {
                l.FieldRefs.Remove(frToRemove);
            }


            foreach (var rf in removeFieldRefAndAddDirectly)
            {
                var lf = spList.Fields.FirstOrDefault(f => f.InternalName == rf);
                if (lf != null)
                {
                    Context.Load(lf, f => f.SchemaXml);
                    Context.ExecuteQuery();

                    l.Fields.Add(new PnP.Framework.Provisioning.Model.Field()
                    {
                        SchemaXml = lf.SchemaXml
                    });
                }
            }
        }
        // We have to replace some of keys and values in our XML file in order to be able to restore in a SPO site, this method will do some replacement for fixing possible issues. 
        private Dictionary<string, string> GetReplacements(ListInstance l, List spList)
        {
            Context.Load(spList.Fields, fields => fields.Include(f => f.InternalName, f => f.Title));
            Context.ExecuteQuery();

            var replacements = l.Fields.ToDictionary(f => this.GetPropertyFromXml(f.SchemaXml, "Name"), f => this.GetPropertyFromXml(f.SchemaXml, "DisplayName"));
            var replacementsFieldRefs = l.FieldRefs.ToDictionary(fr => fr.Name, fr => fr.DisplayName);

            foreach (var fr in replacementsFieldRefs)
            {
                if (!replacements.ContainsKey(fr.Key))
                {
                    replacements.Add(fr.Key, fr.Value);
                }
            }

            if (!replacements.ContainsKey("Title"))
            {
                var titleFieldName = spList != null ? spList.Fields.Where(f => f.InternalName == "Title").Select(f => f.Title).FirstOrDefault() : "";
                replacements.Add("Title", titleFieldName ?? "Title");
            }

            return replacements;
        }
        // Validation in columns will make some problem 
        private string RemoveValidationTags(string schemaXml)
        {
            XDocument xdoc = XDocument.Parse(schemaXml);
            var validationElement = xdoc.Descendants("Validation").FirstOrDefault();

            if (validationElement != null)
            {
                validationElement.Remove();
            }

            return xdoc.ToString();
        }
        // This method will remove calculated validation and calculated value of a columns in our XML file
        private void FixCalculatedColumns(PnP.Framework.Provisioning.Model.Field f, Dictionary<string, string> replacements)
        {
            foreach (var r in replacements)
            {
                if (!string.IsNullOrEmpty(r.Key) && !string.IsNullOrEmpty(r.Value))
                {
                    f.SchemaXml = f.SchemaXml.Replace($"{{fieldtitle:{r.Key}}}", r.Value);
                }
            }
            f.SchemaXml = RemoveValidationTags(f.SchemaXml);
        }
        // Our backup is in XMl format and this method will provide us attribute and value of an xml tags. 
        private string GetPropertyFromXml(string xml, string key)
        {
            var xDoc = XDocument.Parse(xml);
            var attr = xDoc.Root.Attribute(key);
            return attr != null ? attr.Value : "";
        }
        // We will ensure there is / at the end of a url 
        public static string EnsureTrailingSlash(string url)
        {
            if (!string.IsNullOrEmpty(url) && url[url.Length - 1] != '/')
            {
                return url + "/";
            }

            return url;
        }
        private static void EnsureFolderPath(Web web, string fileUrl)
        {
            var webRelativeUrlFolder = fileUrl.Substring(web.ServerRelativeUrl.Length);
            if (webRelativeUrlFolder.Contains("/")) webRelativeUrlFolder = webRelativeUrlFolder.Substring(0, webRelativeUrlFolder.LastIndexOf("/"));
            web.EnsureFolderPath(webRelativeUrlFolder);
        }
        // Log progress of process into console
        private void LogProcessToConsole(String message, Int32 progress, Int32 total)
        {
            // Only to output progress for console UI
            LogToConsole(string.Format("{0:00}/{1:00} - {2}", progress, total, message));
        }
        // Log detailes into console
        private void LogMessageToConsole(string message, ProvisioningMessageType messageType)
        {
            LogToConsole(message);
        }
        // Log to console
        private static void LogToConsole(string message)
        {
            Console.WriteLine("{0}", message);
        }




    }
}
