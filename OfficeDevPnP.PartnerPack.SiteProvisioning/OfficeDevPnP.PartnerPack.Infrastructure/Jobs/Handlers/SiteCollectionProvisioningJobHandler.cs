using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.Online.SharePoint.TenantManagement;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure.Jobs.Handlers
{
    /// <summary>
    /// [Rise]: This class has been extensively modified to handle Rise elements in the PnP Partner Pack
    /// </summary>    
    public class SiteCollectionProvisioningJobHandler : ProvisioningJobHandler
    {
        private SiteCollectionProvisioningJob job;
        private string tenantId;
        private String siteUrl;
        private bool scriptsDisabled;

        protected override void RunJobInternal(ProvisioningJob provisioningJob)
        {
            SiteCollectionProvisioningJob scj = provisioningJob as SiteCollectionProvisioningJob;
            if (scj == null)
            {
                throw new ArgumentException("Invalid job type for SiteCollectionProvisioningJobHandler.");
            }

            // set private members
            job = scj;
            tenantId = job.TenantId;
            siteUrl = GetSiteUrl(job);

            // Init
            CreateSiteCollection();
        }

        private string GetSiteUrl(SiteCollectionProvisioningJob job)
        {
            if (job.RootUrl != null)
            {
                return job.RootUrl + job.RelativeUrl;

            }
            else
            {
                if(ConfigurationManager.AppSettings["SingleTenant"] == "true")
                {
                    return String.Format("{0}{1}", PnPPartnerPackSettings.InfrastructureSiteUrl.Substring(0, PnPPartnerPackSettings.InfrastructureSiteUrl.IndexOf("sharepoint.com/") + 14), job.RelativeUrl);

                } else
                {
                    return String.Format("{0}{1}", PnPPartnerPackSettings.InfrastructureSiteUrlFromTenantId(tenantId).Substring(0, PnPPartnerPackSettings.InfrastructureSiteUrlFromTenantId(tenantId).IndexOf("sharepoint.com/") + 14), job.RelativeUrl);
                }

            }
        }

        private void CreateSiteCollection()
        {
            // Define the full Site Collection URL

            // Load the template from the source Templates Provider
            if (!String.IsNullOrEmpty(job.TemplatesProviderTypeName))
            {
                ProvisioningTemplate template = null;

                var templatesProvider = PnPPartnerPackSettings.TemplatesProvidersFromTenantId(tenantId)[job.TemplatesProviderTypeName];
                if (templatesProvider != null)
                {
                    template = templatesProvider.GetProvisioningTemplateFromTenantId(job.ProvisioningTemplateUrl, tenantId);
                }

                if (template != null)
                {

                    Console.WriteLine("Checking if Site Collection \"{0}\" exists...", job.RelativeUrl);

                    using (var adminContext = PnPPartnerPackContextProvider.GetAppOnlyTenantLevelClientContextFromTenantId(tenantId))
                    {
                        var tenant = new Tenant(adminContext);

                        if (!tenant.SiteExists(siteUrl))
                        {
                            Console.WriteLine("Creating Site Collection \"{0}\".", job.RelativeUrl);

                            adminContext.RequestTimeout = Timeout.Infinite;

                            // Configure the Site Collection properties
                            SiteEntity newSite = new SiteEntity();
                            newSite.Description = job.Description;
                            newSite.Lcid = (uint)job.Language;
                            newSite.Title = job.SiteTitle;
                            newSite.Url = siteUrl;
                            newSite.SiteOwnerLogin = job.PrimarySiteCollectionAdmin;
                            newSite.StorageMaximumLevel = job.StorageMaximumLevel;
                            newSite.StorageWarningLevel = job.StorageWarningLevel;

                            // Use the BaseSiteTemplate of the template, if any, otherwise 
                            // fallback to the pre-configured site template (i.e. STS#0)
                            newSite.Template = !String.IsNullOrEmpty(template.BaseSiteTemplate) ? template.BaseSiteTemplate : PnPPartnerPackSettings.DefaultSiteTemplate;

                            newSite.TimeZoneId = job.TimeZone;
                            newSite.UserCodeMaximumLevel = job.UserCodeMaximumLevel;
                            newSite.UserCodeWarningLevel = job.UserCodeWarningLevel;

                            // Create the Site Collection and wait for its creation
                            tenant.CreateSiteCollection(newSite, true, true);
                            Console.WriteLine("Site collection \"{0}\" created...", siteUrl);

                            if (ScriptsAreDisabledOnSite())
                            {
                                scriptsDisabled = true;
                                SetScriptsOnSite(true);
                            }

                            ApplyProvisioningTemplate(siteUrl, template);

                        }
                        else
                        {
                            Console.WriteLine("Site Collection \"{0}\" already exists...", job.RelativeUrl);

                            if (ScriptsAreDisabledOnSite())
                            {
                                scriptsDisabled = true;
                                SetScriptsOnSite(true);
                            }

                            ApplyProvisioningTemplate(siteUrl, template);
                        }

                    }
                }
            }
        }

        public bool ScriptsAreDisabledOnSite()
        {
            using (var adminContext = PnPPartnerPackContextProvider.GetAppOnlyTenantLevelClientContextFromTenantId(tenantId))
            {
                var tenant = new Tenant(adminContext);
                var siteProperties = tenant.GetSitePropertiesByUrl(siteUrl, true);
                adminContext.Load(siteProperties);
                adminContext.ExecuteQuery();
                DenyAddAndCustomizePagesStatus status = siteProperties.DenyAddAndCustomizePages;
                if (status == DenyAddAndCustomizePagesStatus.Enabled)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }

        public void SetScriptsOnSite(bool enabled)
        {
            using (var adminContext = PnPPartnerPackContextProvider.GetAppOnlyTenantLevelClientContextFromTenantId(tenantId))
            {
                string log = (enabled) ? "Enabling scripts on site..." : "Disabling scripts on site...";
                Console.WriteLine(log);
                var tenant = new Tenant(adminContext);
                var siteProperties = tenant.GetSitePropertiesByUrl(siteUrl, true);
                adminContext.Load(siteProperties);
                adminContext.ExecuteQuery();
                siteProperties.DenyAddAndCustomizePages = (enabled) ? DenyAddAndCustomizePagesStatus.Disabled : DenyAddAndCustomizePagesStatus.Enabled;
                var result = siteProperties.Update();
                adminContext.Load(result);
                adminContext.ExecuteQuery();
            }
        }

        public void ApplyProvisioningTemplate(string siteUrl, ProvisioningTemplate template)
        {
            // Move to the context of the created Site Collection
            using (ClientContext clientContext = PnPPartnerPackContextProvider.GetAppOnlyClientContextFromTenantId(siteUrl, tenantId))
            {
                clientContext.RequestTimeout = Timeout.Infinite;

                Site site = clientContext.Site;
                Web web = site.RootWeb;
                clientContext.Load(site, s => s.Url);
                clientContext.Load(web, w => w.Url);
                clientContext.ExecuteQueryRetry();

                // Apply the Provisioning Template
                Console.WriteLine("Applying Provisioning Template \"{0}\" to site...", job.ProvisioningTemplateUrl);

                // We do intentionally remove taxonomies, which are not supported 
                // in the AppOnly Authorization model
                // For further details, see the PnP Partner Pack documentation 
                ProvisioningTemplateApplyingInformation ptai = new ProvisioningTemplateApplyingInformation();

                // Write provisioning steps on console log
                ptai.MessagesDelegate = (message, type) =>
                {
                    switch (type)
                    {
                        case ProvisioningMessageType.Warning:
                            {
                                Console.WriteLine("{0} - {1}", type, message);
                                break;
                            }
                        case ProvisioningMessageType.Progress:
                            {
                                var activity = message;
                                if (message.IndexOf("|") > -1)
                                {
                                    var messageSplitted = message.Split('|');
                                    if (messageSplitted.Length == 4)
                                    {
                                        var status = messageSplitted[0];
                                        var statusDescription = messageSplitted[1];
                                        var current = double.Parse(messageSplitted[2]);
                                        var total = double.Parse(messageSplitted[3]);
                                        var percentage = Convert.ToInt32((100 / total) * current);
                                        Console.WriteLine("{0} - {1} - {2}", percentage, status, statusDescription);
                                    }
                                    else
                                    {
                                        Console.WriteLine(activity);
                                    }
                                }
                                else
                                {
                                    Console.WriteLine(activity);
                                }
                                break;
                            }
                        case ProvisioningMessageType.Completed:
                            {
                                Console.WriteLine(type);
                                break;
                            }
                    }
                };
                ptai.ProgressDelegate = (message, step, total) =>
                {
                    var percentage = Convert.ToInt32((100 / Convert.ToDouble(total)) * Convert.ToDouble(step));
                    Console.WriteLine("{0:00}/{1:00} - {2} - {3}", step, total, percentage, message);
                };

                // Exclude handlers not supported in App-Only
                ptai.HandlersToProcess ^= OfficeDevPnP.Core.Framework.Provisioning.Model.Handlers.TermGroups;
                ptai.HandlersToProcess ^= OfficeDevPnP.Core.Framework.Provisioning.Model.Handlers.SearchSettings;

                // Configure template parameters
                if (job.TemplateParameters != null)
                {
                    foreach (var key in job.TemplateParameters.Keys)
                    {
                        if (job.TemplateParameters.ContainsKey(key))
                        {
                            template.Parameters[key] = job.TemplateParameters[key];
                        }
                    }
                }

                // Fixup Title and Description
                if (template.WebSettings != null)
                {
                    template.WebSettings.Title = job.SiteTitle;
                    template.WebSettings.Description = job.Description;
                }

                // Replace existing Structural Current Navigation on target site
                if (template.Navigation != null &&
                    template.Navigation.CurrentNavigation != null &&
                    template.Navigation.CurrentNavigation.StructuralNavigation != null &&
                    (template.Navigation.CurrentNavigation.NavigationType == CurrentNavigationType.Structural ||
                    template.Navigation.CurrentNavigation.NavigationType == CurrentNavigationType.StructuralLocal))
                {
                    template.Navigation.CurrentNavigation.StructuralNavigation.RemoveExistingNodes = true;
                }
                else if (template.Navigation != null &&
                    template.Navigation.CurrentNavigation != null &&
                    template.Navigation.CurrentNavigation.ManagedNavigation != null &&
                    template.Navigation.CurrentNavigation.NavigationType == CurrentNavigationType.Managed)
                {
                    // We intentionally skip the Managed Navigation
                    template.Navigation = new Core.Framework.Provisioning.Model.Navigation(template.Navigation.GlobalNavigation, null);
                }

                // Replace existing Structural Global Navigation on target site
                if (template.Navigation != null &&
                    template.Navigation.GlobalNavigation != null &&
                    template.Navigation.GlobalNavigation.StructuralNavigation != null &&
                    template.Navigation.GlobalNavigation.NavigationType == GlobalNavigationType.Structural)
                {
                    template.Navigation.GlobalNavigation.StructuralNavigation.RemoveExistingNodes = true;
                }
                else if (template.Navigation != null &&
                    template.Navigation.GlobalNavigation != null &&
                    template.Navigation.GlobalNavigation.ManagedNavigation != null &&
                    template.Navigation.GlobalNavigation.NavigationType == GlobalNavigationType.Managed)
                {
                    // We intentionally skip the Managed Navigation
                    template.Navigation = new Core.Framework.Provisioning.Model.Navigation(null, template.Navigation.CurrentNavigation);
                }

                // Apply the template to the target site
                web.ApplyProvisioningTemplate(template, ptai);

                // Save the template information in the target site
                var info = new SiteTemplateInfo()
                {
                    TemplateProviderType = job.TemplatesProviderTypeName,
                    TemplateUri = job.ProvisioningTemplateUrl,
                    TemplateParameters = template.Parameters,
                    AppliedOn = DateTime.Now,
                };

                // Set site policy template
                if (!String.IsNullOrEmpty(job.SitePolicy))
                {
                    web.ApplySitePolicy(job.SitePolicy);
                }

                Console.WriteLine("Applied Provisioning Template \"{0}\" to site.", job.ProvisioningTemplateUrl);
                WriteCoreMetadata(web);

                // Do we have a modern team site? If we do, then write the graph extension too...
                if (template.BaseSiteTemplate == "GROUP#0")
                {
                    WriteGraphExtension(web.AllProperties["GroupId"].ToString(), job.CoreMetadata);
                }

                // Update Global.json.js file with menu references for our site...
                if (job.MainMenu != null)
                {
                    WriteRiseMenuReference("Global.json.js", web.Id.ToString(), job.MainMenu, "MenuMapping");
                }

                if (job.FooterMenu != null)
                {
                    WriteRiseMenuReference("Global.json.js", web.Id.ToString(), job.FooterMenu, "FooterMapping");
                }

                // scripts were originally disabled on the site. Set back to this policy after we are done...
                if (scriptsDisabled)
                {
                    SetScriptsOnSite(false);
                }

            }

        }

        public void WriteCoreMetadata(Web web)
        {
            Console.WriteLine("Writing core metadata in site...");
            JObject obj = JObject.Parse(job.CoreMetadata);
            Dictionary<string, object> entries = obj.ToObject<Dictionary<string, object>>();
            foreach (var entry in entries)
            {
                dynamic items = entry.Value;
                string termCollection = "";
                foreach (var item in items)
                {
                    termCollection += "#" + item["Label"] + "|" + item["Id"] + ";";
                }

                string key = "PRFT" + entry.Key;
                web.SetPropertyBagValue(key, termCollection);
                WebExtensions.AddIndexedPropertyBagKey(web, key);
            }
        }

        public void WriteRiseMenuReference(string fileName, string siteId, string menuId, string key)
        {
            Console.WriteLine("Updating " + key + " data on Rise Settings site with site ID: " + siteId + " and menu ID: " + menuId + "...");
            string url = PnPPartnerPackSettings.InfrastructureSiteUrlFromTenantId(tenantId);
            using (ClientContext clientContext = PnPPartnerPackContextProvider.GetAppOnlyClientContextFromTenantId(url, tenantId))
            {
                clientContext.RequestTimeout = Timeout.Infinite;
                Site site = clientContext.Site;
                Web web = site.RootWeb;
                clientContext.Load(site, s => s.Url);
                clientContext.Load(web, w => w.Url);
                var list = clientContext.Web.Lists.GetByTitle("riseData");
                clientContext.Load(list);
                clientContext.ExecuteQuery();
                string webUrl = list.ParentWebUrl;
                Microsoft.SharePoint.Client.File file = web.GetFileByServerRelativeUrl(webUrl + "/riseData/" + fileName);
                clientContext.Load(file);
                clientContext.ExecuteQuery();
                ClientResult<Stream> streamResult = file.OpenBinaryStream();
                clientContext.ExecuteQuery();

                using (MemoryStream mStream = new MemoryStream())
                {
                    if (streamResult != null)
                    {
                        streamResult.Value.CopyTo(mStream);
                        string data = Encoding.UTF8.GetString(mStream.ToArray());
                        string originalData = data;
                        data = data.Replace("window.lokiFiles['Global'] = ", "");
                        dynamic json = JsonConvert.DeserializeObject(data);
                        int index = 0;
                        bool mappingExists = false;

                        foreach (var val in json["collections"])
                        {
                            if (val["name"] == key)
                            {
                                json["collections"][index]["data"].Add(JsonConvert.DeserializeObject("{ \"siteId\": \"" + siteId + "\", \"menuId\": \"" + menuId + "\", \"meta\": { \"revision\": 0, \"created\": " + (long)(DateTime.UtcNow - new DateTime(1970, 1, 1)).TotalMilliseconds + ", \"version\": 0 }, \"$loki\": 1 }"));
                                mappingExists = true;
                            }
                            index++;
                        }

                        // If we got to the end of the collection and there were no mapping entries, create a new Loki collection and add the mapping entry under the key
                        if (!mappingExists)
                        {
                            json["collections"].Add(JsonConvert.DeserializeObject("{\"name\":\"" + key + "\",\"data\":[{ \"siteId\": \"" + siteId + "\", \"menuId\": \"" + menuId + "\", \"meta\": { \"revision\": 0, \"created\": " + (long)(DateTime.UtcNow - new DateTime(1970, 1, 1)).TotalMilliseconds + ", \"version\": 0 }, \"$loki\": 1 }],\"idIndex\":[1],\"binaryIndices\":{},\"constraints\":null,\"uniqueNames\":[],\"transforms\":{},\"objType\":\"MenuMapping\",\"dirty\":false,\"cachedIndex\":null,\"cachedBinaryIndex\":null,\"cachedData\":null,\"adaptiveBinaryIndices\":false,\"transactional\":false,\"cloneObjects\":false,\"cloneMethod\":\"parse-stringify\",\"asyncListeners\":false,\"disableChangesApi\":true,\"autoupdate\":false,\"ttl\":null,\"maxId\":1,\"DynamicViews\":[],\"events\":{\"insert\":[null],\"update\":[null],\"pre-insert\":[],\"pre-update\":[],\"close\":[],\"flushbuffer\":[],\"error\":[],\"delete\":[null],\"warning\":[null]},\"changes\":[]}"));

                        }

                        // Create a backup Loki file in case we screw it up
                        Console.WriteLine("Creating backup Loki file...");
                        FileCreationInformation backUpFile = new FileCreationInformation();
                        backUpFile.Overwrite = true;
                        backUpFile.Url = fileName + "_provisioning_backup";
                        byte[] originalBytes = Encoding.UTF8.GetBytes(originalData);
                        backUpFile.Content = originalBytes;
                        Microsoft.SharePoint.Client.File addedBackupFile = list.RootFolder.Files.Add(backUpFile);

                        // Write the new file
                        Console.WriteLine("Writing new Loki file...");
                        FileCreationInformation newFile = new FileCreationInformation();
                        newFile.Overwrite = true;
                        newFile.Url = fileName;
                        byte[] toBytes = Encoding.UTF8.GetBytes("window.lokiFiles['Global'] = " + Convert.ToString(json));
                        newFile.Content = toBytes;
                        Microsoft.SharePoint.Client.File addedNewFile = list.RootFolder.Files.Add(newFile);
                        clientContext.Load(addedNewFile);
                        clientContext.ExecuteQuery();
                    }
                }
            }
        }

        public void WriteGraphExtension(string groupId, string data)
        {
            Console.WriteLine("Writing graph extension to modern team site...");
            string tokenResponse = GetAppOnlyAccessToken(tenantId);
            dynamic token = JsonConvert.DeserializeObject(tokenResponse);
            var request = (HttpWebRequest)WebRequest.Create("https://graph.microsoft.com/v1.0/groups/" + groupId + "/extensions");
            request.ContentType = "application/json";
            request.Method = "POST";
            request.Headers["Authorization"] = "Bearer " + token["access_token"];

            using (var streamWriter = new StreamWriter(request.GetRequestStream()))
            {
                string json = "{" +
                    "\"@odata.type\":\"microsoft.graph.openTypeExtension\"," +
                    "\"extensionName\":\"PRFTCoreMetadata\"," +
                    "\"data\":" + data + "" +
                    "}";

                streamWriter.Write(json);
                streamWriter.Flush();
                streamWriter.Close();

                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                {
                    using (Stream stream = response.GetResponseStream())
                    {
                        using (StreamReader reader = new StreamReader(stream))
                        {
                            var result = reader.ReadToEnd();
                        }
                    }
                }

            }
        }

        private string GetAppOnlyAccessToken(string tenantId)
        {
            string url = "https://login.microsoftonline.com/" + tenantId + "/oauth2/v2.0/token";
            var request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "POST";
            request.Accept = "application/json; odata=verbose";
            request.ContentType = "application/x-www-form-urlencoded";
            string postData = "client_id=" + ConfigurationManager.AppSettings["ida:ClientId"] + "&scope=https://graph.microsoft.com/.default&client_secret=" + ConfigurationManager.AppSettings["ida:ClientSecret"] + "&grant_type=client_credentials";
            byte[] postArray = System.Text.Encoding.UTF8.GetBytes(postData);
            request.ContentLength = postArray.Length;
            Stream requestStream = request.GetRequestStream();
            requestStream.Write(postArray, 0, postArray.Length);
            requestStream.Close();

            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            {
                using (Stream stream = response.GetResponseStream())
                {
                    using (StreamReader reader = new StreamReader(stream))
                    {
                        var result = reader.ReadToEnd();
                        return result;
                    }
                }
            }

        }
    }
}
