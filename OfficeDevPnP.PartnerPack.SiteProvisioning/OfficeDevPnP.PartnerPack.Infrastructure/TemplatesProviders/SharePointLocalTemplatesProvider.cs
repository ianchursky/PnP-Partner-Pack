using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure.TemplatesProviders
{
    /// <summary>
    /// Implements the Templates Provider that use the local Site Collection as a repository
    /// </summary>
    public class SharePointLocalTemplatesProvider : SharePointBaseTemplatesProvider
    {
        public override string DisplayName
        {
            get { return ("Current Site Collection"); }
        }

        // [Rise]: SharePointLocalTemplatesProvider() and SharePointLocalTemplatesProvider(String tenantId) are additional constructors added.
        public SharePointLocalTemplatesProvider() : base(PnPPartnerPackSettings.ParentSiteUrl)
        {

        }

        public SharePointLocalTemplatesProvider(String tenantId) : base(PnPPartnerPackSettings.ParentSiteUrl, tenantId)
        {

        }

        // NOTE: Use the current context to determine the URL of the current Site Collection

        public override ProvisioningTemplate GetProvisioningTemplate(string templateUri)
        {
            return (base.GetProvisioningTemplate(templateUri));
        }

        // [Rise]: Additional method...
        public override ProvisioningTemplate GetProvisioningTemplateFromTenantId(string templateUri, string tenantId)
        {
            return (base.GetProvisioningTemplateFromTenantId(templateUri, tenantId));
        }

        public override ProvisioningTemplateInformation[] SearchProvisioningTemplates(string searchText, TargetPlatform platforms, TargetScope scope)
        {
            return (base.SearchProvisioningTemplates(searchText, platforms, scope));
        }

    }
}
