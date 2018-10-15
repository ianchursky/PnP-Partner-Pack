﻿using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.PartnerPack.Infrastructure.SharePoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure.TemplatesProviders
{
    /// <summary>
    /// Implements the Templates Provider that use the local tenant-scoped Site Collection as a repository
    /// </summary>
    public class SharePointGlobalTemplatesProvider : SharePointBaseTemplatesProvider
    {
        public override string DisplayName
        {
            get { return ("Global Tenant"); }
        }

        // [Rise]: SharePointGlobalTemplatesProvider() and SharePointGlobalTemplatesProvider(String tenantId) are additional constructors added.
        public SharePointGlobalTemplatesProvider(): base(PnPPartnerPackSettings.InfrastructureSiteUrl)
        {

        }

        public SharePointGlobalTemplatesProvider(String tenantId) : base(PnPPartnerPackSettings.InfrastructureSiteUrl, tenantId)
        {

        }

        public override ProvisioningTemplate GetProvisioningTemplate(string templateUri)
        {
            return (base.GetProvisioningTemplate(templateUri));
        }

        // [Rise]: Additional method...
        public override ProvisioningTemplate GetProvisioningTemplateFromTenantId(string templateUri, string tenantId)
        {
            return (base.GetProvisioningTemplateFromTenantId(templateUri, tenantId));
        }

        public override  ProvisioningTemplateInformation[] SearchProvisioningTemplates(string searchText, TargetPlatform platforms, TargetScope scope)
        {
            return (base.SearchProvisioningTemplates(searchText, platforms, scope));
        }
    }
}
