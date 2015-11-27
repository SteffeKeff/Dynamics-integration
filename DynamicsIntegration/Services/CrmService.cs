using System;
using System.Linq;
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
        string userName, password, domain = "", organizationUniqueName = "", discoveryServiceAddress = " https://disco.crm4.dynamics.com/XRMServices/2011/Discovery.svc";
        OrganizationDetailCollection organizations;
        OrganizationServiceProxy organizationServiceProxy = null;
        #endregion Class Level Members

        public CrmService(DynamicsCredentials credentials) : this(credentials, "") {}

        public CrmService(DynamicsCredentials credentials, string orgName) 
        {
            userName = credentials.UserName;
            password = credentials.Password;
            domain = credentials.Domain;
            organizationUniqueName = orgName;

            organizationServiceProxy = getOrganizationServiceProxy();
        }

        public void logout()
        {
            organizationServiceProxy.Dispose();
        }

        public Dictionary<string, string> getOrganizations()
        {
            Dictionary<string, string> orgs = new Dictionary<string, string>();

            foreach (var details in organizations.ToArray())
            {
                orgs.Add(details.FriendlyName, details.UniqueName);
            }

            return orgs;
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

        public EntityCollection getContactsInList(string id, bool allAttributes, int top)
        {
            EntityCollection listMembers, contacts;
            Guid listid = new Guid(id);
            QueryExpression contactsQuery;

            var listMembersQuery = new QueryExpression { EntityName = "listmember", ColumnSet = new ColumnSet("listid", "entityid") };
            
            if (allAttributes)
            {
                contactsQuery = new QueryExpression { EntityName = "contact", ColumnSet = new ColumnSet(true) };
            }
            else
            {
                contactsQuery = new QueryExpression { EntityName = "contact", ColumnSet = new ColumnSet("firstname", "lastname", "emailaddress1", "mobilephone") };
            }

            if (top != 0){ listMembersQuery.TopCount = top; }

            listMembersQuery.Criteria = new FilterExpression();
            listMembersQuery.Criteria.AddCondition("listid", ConditionOperator.Equal, listid);

            listMembers = organizationServiceProxy.RetrieveMultiple(listMembersQuery);

            contactsQuery.Criteria = new FilterExpression();
            ConditionExpression condition = new ConditionExpression();
            condition.AttributeName = "contactid";
            condition.Operator = ConditionOperator.In;

            foreach(ListMember member in listMembers.Entities)
            {
                condition.Values.Add(member.EntityId.Id);
            }

            contactsQuery.Criteria.AddCondition(condition);

            contacts = organizationServiceProxy.RetrieveMultiple(contactsQuery);

            return contacts;
        }

        public Dictionary<string, string> GetAttributeDisplayName(string entitySchemaName)
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

        private string getOrgName()
        {
            if(String.IsNullOrEmpty(organizationUniqueName))
            {
                return organizations.ToArray()[0].UniqueName;
            }
            return organizationUniqueName;
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
                    organizations = DiscoverOrganizations(discoveryProxy);
                    organizationUniqueName = String.IsNullOrEmpty(organizationUniqueName) ? organizations.ToArray()[0].UniqueName : organizationUniqueName;
                    
                    organizationUri = FindOrganization(organizationUniqueName,
                         organizations.ToArray()).Endpoints[EndpointType.OrganizationService];
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
