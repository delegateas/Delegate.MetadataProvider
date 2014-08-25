using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.Serialization;
using System.ServiceModel.Description;
using Microsoft.Crm.Services.Utility;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Metadata;

namespace DG.MetadataProvider
{
    public sealed class IfdMetadataProviderService : IMetadataProviderService
    {
        private class OrganizationMetadata : IOrganizationMetadata
        {
            public EntityMetadata[] Entities { get; set; }
            public OptionSetMetadataBase[] OptionSets { get; set; }
            public SdkMessages Messages { get; set; }
        }

        private readonly Dictionary<string, string> _parameters;
        private OrganizationMetadata _metadata;

        public IfdMetadataProviderService(IDictionary<string, string> parameters)
        {
            _parameters = new Dictionary<string, string>(parameters, StringComparer.InvariantCultureIgnoreCase);
        }

        IOrganizationMetadata IMetadataProviderService.LoadMetadata()
        {
            if (_metadata == null)
            {
                var credentials = new ClientCredentials();
                credentials.UserName.UserName = _parameters["username"];
                credentials.UserName.Password = _parameters["password"];
                using (var service = new OrganizationServiceProxy(new Uri(_parameters["url"]), null, credentials, null))
                {
                    service.Timeout = new TimeSpan(0, 5, 0);
                    _metadata = new OrganizationMetadata
                    {
                        Entities = RetrieveEntities(service),
                        OptionSets = RetrieveOptionSets(service),
                        Messages = RetrieveMessages(service),
                    };
                }
            }
            return _metadata;
        }

        private static EntityMetadata[] RetrieveEntities(IOrganizationService service)
        {
            var request = new OrganizationRequest("RetrieveAllEntities");
            request.Parameters["EntityFilters"] = EntityFilters.Entity | EntityFilters.Attributes | EntityFilters.Relationships;
            request.Parameters["RetrieveAsIfPublished"] = false;
            var response = service.Execute(request);
            return response.Results["EntityMetadata"] as EntityMetadata[];
        }

        private static OptionSetMetadataBase[] RetrieveOptionSets(IOrganizationService service)
        {
            var request = new OrganizationRequest("RetrieveAllOptionSets");
            request.Parameters["RetrieveAsIfPublished"] = false;
            var response = service.Execute(request);
            return response.Results["OptionSetMetadata"] as OptionSetMetadataBase[];
        }

        private static SdkMessages RetrieveMessages(IOrganizationService service)
        {
            var type = Assembly.LoadFrom("crmsvcutil.exe").GetType("Microsoft.Crm.Services.Utility.SdkMetadataProviderService");
            var getMessages = type.GetMethod("RetrieveSdkRequests", BindingFlags.Instance | BindingFlags.NonPublic);
            var provider = FormatterServices.GetUninitializedObject(type);
            return getMessages.Invoke(provider, new object[] { service }) as SdkMessages;
        }
    }
}