// =====================================================================
//  This file is part of the Microsoft Dynamics CRM SDK code samples.
//
//  Copyright (C) Microsoft Corporation.  All rights reserved.
//
//  This source code is intended only as a supplement to Microsoft
//  Development Tools and/or on-line documentation.  See these other
//  materials for detailed information regarding Microsoft code samples.
//
//  THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
//  KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
//  IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
//  PARTICULAR PURPOSE.
// =====================================================================

//<snippetAuthenticateWithNoHelp>
using System;
using System.ServiceModel.Description;
using System.Collections;
using System.Collections.Generic;

// These namespaces are found in the Microsoft.Xrm.Sdk.dll assembly
// located in the SDK\bin folder of the SDK download.
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Discovery;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

using System.Linq;
using DynamicsIntegration.Models;

namespace DynamicsIntegration.Controllers
{
    /// <summary>
    /// Demonstrate how to do basic authentication using IServiceManagement and SecurityTokenResponse.
    /// </summary>
    class CrmService
    {
        #region Class Level Members
        // To get discovery service address and organization unique name, 
        // Sign in to your CRM org and click Settings, Customization, Developer Resources.
        // On Developer Resource page, find the discovery service address under Service Endpoints and organization unique name under Your Organization Information.
        //private String _discoveryServiceAddress = "https://dev.crm.dynamics.com/XRMServices/2011/Discovery.svc";
        private String _discoveryServiceAddress = " https://disco.crm4.dynamics.com/XRMServices/2011/Discovery.svc";
        //private String _organizationUniqueName = "OrganizationUniqueName";
        //private String _organizationUniqueName = "org66aa9e14";
        // Provide your user name and password.
        public string _userName;
        public String _password;

        // Provide domain name for the On-Premises org.
        public String _domain = "";
        private OrganizationServiceProxy organizationServiceProxy = null;

        #endregion Class Level Members

        public CrmService(DynamicsCredentials credentials)
        {
            _userName = credentials.UserName;
            _password = credentials.Password;
            _domain = credentials.Domain;

            organizationServiceProxy = getOrganizationServiceProxy();
        }

        public EntityCollection getAllLists(bool allAttributes)
        {
            EntityCollection results = new EntityCollection();
            QueryExpression query;

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

        public string changeBulkEmail(string contactToUpdate)
        {
            
            Contact contact = organizationServiceProxy.Retrieve("contact", new Guid(contactToUpdate), new ColumnSet("donotbulkemail")).ToEntity<Contact>();

            if (contact.DoNotBulkEMail == true)
            {
                contact.DoNotBulkEMail = false;
            }
            else
            {
                contact.DoNotBulkEMail = true;
            }

            organizationServiceProxy.Update(contact);

            return "Contact successfully updated!";
        }

        public ArrayList getContactsInList(string id, bool allAttributes, int preview)
        {
            ArrayList contacts = new ArrayList();
            Guid listid;
            EntityCollection results = new EntityCollection();
            QueryExpression query = new QueryExpression { EntityName = "listmember", ColumnSet = new ColumnSet("listid", "entityid") };

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

            return contacts;
        }

        public Dictionary<string,string> GetAttributeDisplayName(string entitySchemaName) /*, string attributeSchemaName*/
        {

            IOrganizationService service = organizationServiceProxy;
            /*RetrieveAttributeRequest retrieveAttributeRequest = new RetrieveAttributeRequest
            {
                EntityLogicalName = entitySchemaName,
                LogicalName = attributeSchemaName,
                RetrieveAsIfPublished = true
            };*/
            RetrieveEntityRequest req = new RetrieveEntityRequest();
            req.RetrieveAsIfPublished = true;
            req.LogicalName = entitySchemaName;
            req.EntityFilters = EntityFilters.Attributes;

            //RetrieveAttributeResponse retrieveAttributeResponse = (RetrieveAttributeResponse)service.Execute(req);//(retrieveAttributeRequest);
            RetrieveEntityResponse resp = (RetrieveEntityResponse)service.Execute(req);

            Dictionary<string, string> displayNames = new Dictionary<string, string>();
             
            for (int iCnt = 0; iCnt < resp.EntityMetadata.Attributes.ToList().Count; iCnt++)
            {
                if (resp.EntityMetadata.Attributes.ToList()[iCnt].DisplayName.LocalizedLabels.Count > 0)
                {
                    string displayName = resp.EntityMetadata.Attributes.ToList()[iCnt].DisplayName.LocalizedLabels[0].Label;
                    string logicalName = resp.EntityMetadata.Attributes.ToList()[iCnt].LogicalName;
                    displayNames.Add(logicalName, displayName.ToString());
                }

            }

            return displayNames;

            //AttributeMetadata retrievedAttributeMetadata = (AttributeMetadata)retrieveAttributeResponse.AttributeMetadata;

           // return retrievedAttributeMetadata.DisplayName.UserLocalizedLabel.Label;
        }


        public OrganizationServiceProxy getOrganizationServiceProxy()
        {
            //<snippetAuthenticateWithNoHelp1>
            IServiceManagement<IDiscoveryService> serviceManagement =
                        ServiceConfigurationFactory.CreateManagement<IDiscoveryService>(
                        new Uri(_discoveryServiceAddress));
            AuthenticationProviderType endpointType = serviceManagement.AuthenticationType;

            // Set the credentials.
            AuthenticationCredentials authCredentials = GetCredentials(serviceManagement, endpointType);

            String organizationUri = String.Empty;
            // Get the discovery service proxy.
            using (DiscoveryServiceProxy discoveryProxy =
                GetProxy<IDiscoveryService, DiscoveryServiceProxy>(serviceManagement, authCredentials))
            {
                // Obtain organization information from the Discovery service. 
                if (discoveryProxy != null)
                {
                    // Obtain information about the organizations that the system user belongs to.
                    OrganizationDetailCollection orgs = DiscoverOrganizations(discoveryProxy);

                    //Fetches the first uniqueName in organizations array
                    string _organizationUniqueName = orgs.ToArray()[0].UniqueName;

                    // Obtains the Web address (Uri) of the target organization.
                    organizationUri = FindOrganization(_organizationUniqueName,
                         orgs.ToArray()).Endpoints[EndpointType.OrganizationService];

                }
            }
            //</snippetAuthenticateWithNoHelp1>

            if (!String.IsNullOrWhiteSpace(organizationUri))
            {
                //<snippetAuthenticateWithNoHelp3>
                IServiceManagement<IOrganizationService> orgServiceManagement =
                    ServiceConfigurationFactory.CreateManagement<IOrganizationService>(
                    new Uri(organizationUri));

                // Set the credentials.
                AuthenticationCredentials credentials = GetCredentials(orgServiceManagement, endpointType);

                // Get the organization service proxy.
                using (organizationServiceProxy =
                    GetProxy<IOrganizationService, OrganizationServiceProxy>(orgServiceManagement, credentials))
                {
                    // This statement is required to enable early-bound type support.
                    organizationServiceProxy.EnableProxyTypes();

                }
            }

            return organizationServiceProxy;
        }


        //<snippetAuthenticateWithNoHelp2>
        /// <summary>
        /// Obtain the AuthenticationCredentials based on AuthenticationProviderType.
        /// </summary>
        /// <param name="service">A service management object.</param>
        /// <param name="endpointType">An AuthenticationProviderType of the CRM environment.</param>
        /// <returns>Get filled credentials.</returns>
        private AuthenticationCredentials GetCredentials<TService>(IServiceManagement<TService> service, AuthenticationProviderType endpointType)
        {
            AuthenticationCredentials authCredentials = new AuthenticationCredentials();

            switch (endpointType)
            {
                case AuthenticationProviderType.ActiveDirectory:
                    authCredentials.ClientCredentials.Windows.ClientCredential =
                        new System.Net.NetworkCredential(_userName,
                            _password,
                            _domain);
                    break;
                case AuthenticationProviderType.LiveId:
                    authCredentials.ClientCredentials.UserName.UserName = _userName;
                    authCredentials.ClientCredentials.UserName.Password = _password;
                    authCredentials.SupportingCredentials = new AuthenticationCredentials();
                    authCredentials.SupportingCredentials.ClientCredentials =
                        Microsoft.Crm.Services.Utility.DeviceIdManager.LoadOrRegisterDevice();
                    break;
                default: // For Federated and OnlineFederated environments.                    
                    authCredentials.ClientCredentials.UserName.UserName = _userName;
                    authCredentials.ClientCredentials.UserName.Password = _password;
                    // For OnlineFederated single-sign on, you could just use current UserPrincipalName instead of passing user name and password.
                    // authCredentials.UserPrincipalName = UserPrincipal.Current.UserPrincipalName;  // Windows Kerberos

                    // The service is configured for User Id authentication, but the user might provide Microsoft
                    // account credentials. If so, the supporting credentials must contain the device credentials.
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
        //</snippetAuthenticateWithNoHelp2>

        /// <summary>
        /// Discovers the organizations that the calling user belongs to.
        /// </summary>
        /// <param name="service">A Discovery service proxy instance.</param>
        /// <returns>Array containing detailed information on each organization that 
        /// the user belongs to.</returns>
        public OrganizationDetailCollection DiscoverOrganizations(
            IDiscoveryService service)
        {
            if (service == null) throw new ArgumentNullException("service");
            RetrieveOrganizationsRequest orgRequest = new RetrieveOrganizationsRequest();
            RetrieveOrganizationsResponse orgResponse =
                (RetrieveOrganizationsResponse)service.Execute(orgRequest);

            return orgResponse.Details;
        }

        /// <summary>
        /// Finds a specific organization detail in the array of organization details
        /// returned from the Discovery service.
        /// </summary>
        /// <param name="orgUniqueName">The unique name of the organization to find.</param>
        /// <param name="orgDetails">Array of organization detail object returned from the discovery service.</param>
        /// <returns>Organization details or null if the organization was not found.</returns>
        /// <seealso cref="DiscoveryOrganizations"/>
        public OrganizationDetail FindOrganization(string orgUniqueName,
            OrganizationDetail[] orgDetails)
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

        /// <summary>
        /// Generic method to obtain discovery/organization service proxy instance.
        /// </summary>
        /// <typeparam name="TService">
        /// Set IDiscoveryService or IOrganizationService type to request respective service proxy instance.
        /// </typeparam>
        /// <typeparam name="TProxy">
        /// Set the return type to either DiscoveryServiceProxy or OrganizationServiceProxy type based on TService type.
        /// </typeparam>
        /// <param name="serviceManagement">An instance of IServiceManagement</param>
        /// <param name="authCredentials">The user's Microsoft Dynamics CRM logon credentials.</param>
        /// <returns></returns>
        /// <snippetAuthenticateWithNoHelp4>
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
                AuthenticationCredentials tokenCredentials =
                    serviceManagement.Authenticate(authCredentials);
                // Obtain discovery/organization service proxy for Federated, LiveId and OnlineFederated environments. 
                // Instantiate a new class of type using the 2 parameter constructor of type IServiceManagement and SecurityTokenResponse.
                return (TProxy)classType
                    .GetConstructor(new Type[] { typeof(IServiceManagement<TService>), typeof(SecurityTokenResponse) })
                    .Invoke(new object[] { serviceManagement, tokenCredentials.SecurityTokenResponse });
            }

            // Obtain discovery/organization service proxy for ActiveDirectory environment.
            // Instantiate a new class of type using the 2 parameter constructor of type IServiceManagement and ClientCredentials.
            return (TProxy)classType
                .GetConstructor(new Type[] { typeof(IServiceManagement<TService>), typeof(ClientCredentials) })
                .Invoke(new object[] { serviceManagement, authCredentials.ClientCredentials });
        }
        /// </snippetAuthenticateWithNoHelp4
    }
}
//</snippetAuthenticateWithNoHelp>