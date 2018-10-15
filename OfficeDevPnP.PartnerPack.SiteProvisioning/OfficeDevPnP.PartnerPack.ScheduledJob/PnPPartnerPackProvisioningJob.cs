using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.TimerJobs;
using OfficeDevPnP.PartnerPack.Infrastructure;
using OfficeDevPnP.PartnerPack.Infrastructure.Jobs;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace OfficeDevPnP.PartnerPack.ScheduledJob
{
    public class PnPPartnerPackProvisioningJob : TimerJob
    {
        public PnPPartnerPackProvisioningJob(): base("PnP Partner Pack Provisioning Job")
        {
            TimerJobRun += ExecuteProvisioningJobs;
        }
        private void ExecuteProvisioningJobs(object sender, TimerJobRunEventArgs e)
        {
            Console.WriteLine("Starting job");

            // Show the current context
            Web web = e.SiteClientContext.Web;
            web.EnsureProperty(w => w.Title);
            
            if (ConfigurationManager.AppSettings["SingleTenant"] == "true")
            {
                Console.WriteLine("App is single tenant version...");

                var provisioningJobs = ProvisioningRepositoryFactory.Current.GetTypedProvisioningJobs<ProvisioningJob>(
                    ProvisioningJobStatus.Pending);

                foreach (var job in provisioningJobs)
                {
                    Console.WriteLine("Processing job: {0} - Owner: {1} - Title: {2}", job.JobId, job.Owner, job.Title);

                    Type jobType = job.GetType();

                    if (PnPPartnerPackSettings.ScheduledJobHandlers.ContainsKey(jobType))
                    {
                        PnPPartnerPackSettings.ScheduledJobHandlers[jobType].RunJob(job);
                    }
                }
            } else
            {
                Console.WriteLine("App is multiple tenant version...");

                Console.WriteLine("Getting Azure tenant IDs from database...");

                HttpWebRequest webRequest = WebRequest.CreateHttp(String.Format("{0}/api/tenant/list", ConfigurationManager.AppSettings["RiseAPIUrl"]));
                using (HttpWebResponse response = (HttpWebResponse)webRequest.GetResponse())
                {
                    using (Stream stream = response.GetResponseStream())
                    {
                        using (StreamReader reader = new StreamReader(stream))
                        {
                            string jsonString = reader.ReadToEnd();
                            JArray tenantIds = JArray.Parse(jsonString);

                            Console.WriteLine("{0} Azure tenant IDs found in database...", tenantIds.Count);

                            foreach (string tenantId in tenantIds)
                            {

                                Console.WriteLine("Processing jobs in Site: {0} (Tenant: {1})...", web.Title, tenantId);

                                // Retrieve the list of pending jobs
                                var provisioningJobs = ProvisioningRepositoryFactory.Current.GetTypedProvisioningJobs<ProvisioningJob>(
                                ProvisioningJobStatus.Pending, tenantId);

                                Console.WriteLine("{0} jobs with \"Pending\" provisioning status found...", provisioningJobs.Length);

                                foreach (var job in provisioningJobs)
                                {
                                    Console.WriteLine("Processing job: {0} - Owner: {1} - Title: {2}...", job.JobId, job.Owner, job.Title);

                                    Type jobType = job.GetType();

                                    if (PnPPartnerPackSettings.ScheduledJobHandlers.ContainsKey(jobType))
                                    {
                                        PnPPartnerPackSettings.ScheduledJobHandlers[jobType].RunJob(job);
                                    }
                                }
                            }
                        }
                    }
                }
            }

            Console.WriteLine("Ending job");
        }
    }
}
