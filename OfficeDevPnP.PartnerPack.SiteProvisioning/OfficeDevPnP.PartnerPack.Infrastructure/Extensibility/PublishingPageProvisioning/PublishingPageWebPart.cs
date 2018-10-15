using OfficeDevPnP.Core.Framework.Provisioning.Model;
namespace OfficeDevPnP.PartnerPack.Infrastructure.Extensibility.PublishingPageProvisioning
{
    public class PublishingPageWebPart : WebPart
    {
        public string DefaultViewDisplayName { get; set; }
        public bool IsListViewWebPart
        {
            get
            {
                return DefaultViewDisplayName != null;
            }

        }
    }
}