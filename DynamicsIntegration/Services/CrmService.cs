using System;
using System.Linq;
using System.Diagnostics;
using System.Collections;
using System.Collections.Generic;
using System.ServiceModel.Description;

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Discovery;

using DynamicsIntegration.Models;

namespace DynamicsIntegration.Controllers
{

    class CrmService
    {
        #region Class Level Members
        //private String organizationUniqueName = "org66aa9e14";
        string userName;
        string password;
        string domain = "";
        string discoveryServiceAddress = " https://disco.crm4.dynamics.com/XRMServices/2011/Discovery.svc";
        OrganizationServiceProxy organizationServiceProxy = null;

        #endregion Class Level Members

        public CrmService(DynamicsCredentials credentials)
        {
            userName = credentials.UserName;
            password = credentials.Password;
            domain = credentials.Domain;

            organizationServiceProxy = getOrganizationServiceProxy();
        }

        public EntityCollection getAllLists(bool allAttributes)
        {
            QueryExpression query;
            var results = new EntityCollection();

            if (allAttributes)
            {
                query = new QueryExpression { EntityName = "list", ColumnSet = new ColumnSet(true) };
            }
            else
            {
                query = new QueryExpression { EntityName = "list", ColumnSet = new ColumnSet("listname", "membercount", "listid", "modifiedon") };
            }
            
            query.AddOrder("modifiedon", OrderType.Descending);

            results = organizationServiceProxy.RetrieveMultiple(query);

            return results;
        }

        public void changeBulkEmail(string contactsToUpdate)
        {   
            var contact = organizationServiceProxy.Retrieve("contact", new Guid(contactsToUpdate), new ColumnSet("donotbulkemail")).ToEntity<Contact>();

            contact.DoNotBulkEMail = contact.DoNotBulkEMail == true ? false : true;

            organizationServiceProxy.Update(contact);
        }

        public ArrayList getContactsInList(string id, bool allAttributes, int preview)
        {
            Debug.WriteLine("getcontactsinlist() " + DateTime.Now.ToString());
            Guid listid;
            var contacts = new ArrayList();
            var results = new EntityCollection();
            var query = new QueryExpression { EntityName = "listmember", ColumnSet = new ColumnSet("listid", "entityid") };

            try
            {
                listid = new Guid(id);
            }
            catch (FormatException)
            {
                return contacts;
            }

            if(preview != 0)
            {
                query.TopCount = preview;
            }

            query.Criteria = new FilterExpression();
            query.Criteria.AddCondition("listid", ConditionOperator.Equal, listid);

            results = organizationServiceProxy.RetrieveMultiple(query);

            foreach (ListMember member in results.Entities)
            {
                Contact contact;
                if (allAttributes)
                {
                    contact = organizationServiceProxy.Retrieve("contact", member.EntityId.Id, new ColumnSet(true)).ToEntity<Contact>();
                }
                else
                {
                    contact = organizationServiceProxy.Retrieve("contact", member.EntityId.Id, new ColumnSet(new string[] { "firstname", "lastname", "emailaddress1", "mobilephone" })).ToEntity<Contact>();
                }
                contacts.Add(contact);
            }
            Debug.WriteLine("klar med getAllLists() klar " + DateTime.Now.ToString());
            return contacts;
        }

        public Dictionary<string,string> GetAttributeDisplayName(string entitySchemaName)
        {
            var service = organizationServiceProxy;
            var req = new RetrieveEntityRequest();
            req.RetrieveAsIfPublished = true;
            req.LogicalName = entitySchemaName;
            req.EntityFilters = EntityFilters.Attributes;

            var resp = (RetrieveEntityResponse)service.Execute(req);
            Dictionary<string, string> displayNames = new Dictionary<string, string>();
             
            for (int iCnt = 0; iCnt < resp.EntityMetadata.Attributes.ToList().Count; iCnt++)
            {
                if (resp.EntityMetadata.Attributes.ToList()[iCnt].DisplayName.LocalizedLabels.Count > 0)
                {
                    var displayName = resp.EntityMetadata.Attributes.ToList()[iCnt].DisplayName.LocalizedLabels[0].Label;
                    var logicalName = resp.EntityMetadata.Attributes.ToList()[iCnt].LogicalName;
                    displayNames.Add(logicalName, displayName.ToString());
                }
            }

            return displayNames;
        }


        public OrganizationServiceProxy getOrganizationServiceProxy()
        {
            IServiceManagement<IDiscoveryService> serviceManagement =
                        ServiceConfigurationFactory.CreateManagement<IDiscoveryService>(
                        new Uri(discoveryServiceAddress));

            var endpointType = serviceManagement.AuthenticationType;
            var authCredentials = GetCredentials(serviceManagement, endpointType);

            var organizationUri = String.Empty;
            using (DiscoveryServiceProxy discoveryProxy =
                GetProxy<IDiscoveryService, DiscoveryServiceProxy>(serviceManagement, authCredentials))
            {
                if (discoveryProxy != null)
                {
                    var orgs = DiscoverOrganizations(discoveryProxy);

                    //Fetches the first uniqueName in organizations array
                    var organizationUniqueName = orgs.ToArray()[0].UniqueName;

                    // Obtains the Web address (Uri) of the target organization.
                    organizationUri = FindOrganization(organizationUniqueName,
                         orgs.ToArray()).Endpoints[EndpointType.OrganizationService];
                }
            }

            if (!String.IsNullOrWhiteSpace(organizationUri))
            {
                IServiceManagement<IOrganizationService> orgServiceManagement =
                    ServiceConfigurationFactory.CreateManagement<IOrganizationService>(
                    new Uri(organizationUri));
                
                var credentials = GetCredentials(orgServiceManagement, endpointType);
                
                using (organizationServiceProxy =
                    GetProxy<IOrganizationService, OrganizationServiceProxy>(orgServiceManagement, credentials))
                {
                    organizationServiceProxy.EnableProxyTypes();
                }
            }

            return organizationServiceProxy;
        }
        
        private AuthenticationCredentials GetCredentials<TService>(IServiceManagement<TService> service, AuthenticationProviderType endpointType)
        {
            var authCredentials = new AuthenticationCredentials();

            switch (endpointType)
            {
                case AuthenticationProviderType.ActiveDirectory:
                    authCredentials.ClientCredentials.Windows.ClientCredential =
                        new System.Net.NetworkCredential(userName,
                            password,
                            domain);
                    break;
                case AuthenticationProviderType.LiveId:
                    authCredentials.ClientCredentials.UserName.UserName = userName;
                    authCredentials.ClientCredentials.UserName.Password = password;
                    authCredentials.SupportingCredentials = new AuthenticationCredentials();
                    authCredentials.SupportingCredentials.ClientCredentials =
                        Microsoft.Crm.Services.Utility.DeviceIdManager.LoadOrRegisterDevice();
                    break;
                default:                   
                    authCredentials.ClientCredentials.UserName.UserName = userName;
                    authCredentials.ClientCredentials.UserName.Password = password;

                    if (endpointType == AuthenticationProviderType.OnlineFederation)
                    {
                        IdentityProvider provider = service.GetIdentityProvider(authCredentials.ClientCredentials.UserName.UserName);
                        if (provider != null && provider.IdentityProviderType == IdentityProviderType.LiveId)
                        {
                            authCredentials.SupportingCredentials = new AuthenticationCredentials();
                            authCredentials.SupportingCredentials.ClientCredentials =
                                Microsoft.Crm.Services.Utility.DeviceIdManager.LoadOrRegisterDevice();
                        }
                    }

                    break;
            }

            return authCredentials;
        }
        
        public OrganizationDetailCollection DiscoverOrganizations(IDiscoveryService service)
        {
            if (service == null) throw new ArgumentNullException("service");
            RetrieveOrganizationsRequest orgRequest = new RetrieveOrganizationsRequest();
            RetrieveOrganizationsResponse orgResponse =
                (RetrieveOrganizationsResponse)service.Execute(orgRequest);

            return orgResponse.Details;
        }
        
        public OrganizationDetail FindOrganization(string orgUniqueName, OrganizationDetail[] orgDetails)
        {
            if (String.IsNullOrWhiteSpace(orgUniqueName))
                throw new ArgumentNullException("orgUniqueName");
            if (orgDetails == null)
                throw new ArgumentNullException("orgDetails");
            OrganizationDetail orgDetail = null;

            foreach (OrganizationDetail detail in orgDetails)
            {
                if (String.Compare(detail.UniqueName, orgUniqueName,
                    StringComparison.InvariantCultureIgnoreCase) == 0)
                {
                    orgDetail = detail;
                    break;
                }
            }
            return orgDetail;
        }
        
        private TProxy GetProxy<TService, TProxy>(
            IServiceManagement<TService> serviceManagement,
            AuthenticationCredentials authCredentials)
            where TService : class
            where TProxy : ServiceProxy<TService>
        {
            Type classType = typeof(TProxy);

            if (serviceManagement.AuthenticationType !=
                AuthenticationProviderType.ActiveDirectory)
            {
                var tokenCredentials =
                    serviceManagement.Authenticate(authCredentials);
                return (TProxy)classType
                    .GetConstructor(new Type[] { typeof(IServiceManagement<TService>), typeof(SecurityTokenResponse) })
                    .Invoke(new object[] { serviceManagement, tokenCredentials.SecurityTokenResponse });
            }
            
            return (TProxy)classType
                .GetConstructor(new Type[] { typeof(IServiceManagement<TService>), typeof(ClientCredentials) })
                .Invoke(new object[] { serviceManagement, authCredentials.ClientCredentials });
        }
    }
}