using System;
using OfficeDevPnP.PartnerPack.Infrastructure;
using System.Net;
using System.IO;
using Newtonsoft.Json.Linq;
using System.Configuration;

namespace OfficeDevPnP.PartnerPack.ScheduledJob
{
        /// <summary>
        /// [Rise]: Scheduled job modified for Rise customers
        /// </summary>    
    class Program
    {
        static void Main()
        {
            var job = new PnPPartnerPackProvisioningJob();
            job.UseThreading = false;

            if (ConfigurationManager.AppSettings["SingleTenant"] == "true")
            {
                job.AddSite(PnPPartnerPackSettings.InfrastructureSiteUrl);

                job.UseAzureADAppOnlyAuthentication(
                    PnPPartnerPackSettings.ClientId,
                    PnPPartnerPackSettings.Tenant,
                    PnPPartnerPackSettings.AppOnlyCertificate);
            }
            else { 
                HttpWebRequest webRequest = WebRequest.CreateHttp(String.Format("{0}/api/tenant/list", ConfigurationManager.AppSettings["RiseAPIUrl"]));
                using (HttpWebResponse response = (HttpWebResponse)webRequest.GetResponse())
                {
                    using (Stream stream = response.GetResponseStream())
                    {
                        using (StreamReader reader = new StreamReader(stream))
                        {
                            string jsonString = reader.ReadToEnd();
                            JArray tenantIds = JArray.Parse(jsonString);
                            foreach (string tenantId in tenantIds)
                            {
                                job.AddSite(PnPPartnerPackSettings.InfrastructureSiteUrlFromTenantId(tenantId));
                            }

                            job.UseAzureADAppOnlyAuthentication(
                                PnPPartnerPackSettings.ClientId,
                                PnPPartnerPackSettings.Tenant,
                                PnPPartnerPackSettings.AppOnlyCertificate);
                        }
                    }
                }
            }
            job.Run();

            #if DEBUG
                Console.WriteLine("Scheduled jobs finished. Press any key to exit...");
                Console.ReadLine();
            #endif
        }
    }
}
