using Microsoft.Xrm.Sdk;

namespace McTools.Xrm.Connection
{
    public class ConnectionDetail
    {
        public CrmServiceClientStub GetCrmServiceClient() => new CrmServiceClientStub();
    }

    // Minimal stub of CrmServiceClient used by the plugin code for compilation only
    public class CrmServiceClientStub
    {
        public System.Guid? CallerId { get; set; }

        public OrganizationResponse Execute(OrganizationRequest request)
        {
            return new OrganizationResponse();
        }
    }
}