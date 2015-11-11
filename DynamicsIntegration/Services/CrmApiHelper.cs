using Microsoft.Xrm.Sdk;
using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using DynamicsIntegration.Controllers;

namespace DynamicsIntegration.Services
{
    public class CrmApiHelper
    {

        CrmService crmService;

        internal CrmApiHelper(CrmService service)
        {
            crmService = service;
        }

        public JObject getValuesFromLists(EntityCollection lists, bool translate, bool allAttributes)
        {
            var responsObject = new JObject();
            var jsonLists = new JArray();
            
            foreach (List list in lists.Entities)
            {
                var jsonList = new JObject();

                if (allAttributes)
                {
                    foreach (KeyValuePair<string, Object> attribute in list.Attributes)
                    {
                        jsonList[attribute.Key] = attribute.Value.ToString();
                    }
                }
                else
                {
                    jsonList["name"] = list.ListName;
                    jsonList["listid"] = list.ListId;
                    jsonList["membercount"] = list.MemberCount;
                    jsonList["modifiedon"] = list.ModifiedOn;
                }

                jsonLists.Add(jsonList);
            }

            if (translate)
            {
                jsonLists = translateToDisplayName(jsonLists, "list");
            }

            responsObject["lists"] = jsonLists;

            return responsObject;
        }

        public JObject getValuesFromContacts(ArrayList contacts, bool translate, bool allAttributes)
        {
            var responsObject = new JObject();
            var jsonContacts = new JArray();

            foreach (Contact contact in contacts)
            {
                var jsonContact = new JObject();

                if (allAttributes)
                {
                    foreach (var prop in contact.GetType().GetProperties())
                    {
                        if (prop.Name != "Item")
                        {
                            var val = prop.GetValue(contact, null);
                            jsonContact[prop.Name] = val == null ? null : val.ToString();
                        }
                    }
                }
                else
                {
                    jsonContact["firstname"] = contact.FirstName;
                    jsonContact["lastname"] = contact.LastName;
                    jsonContact["emailaddress1"] = contact.EMailAddress1;
                    jsonContact["mobilephone"] = contact.MobilePhone;
                }

                jsonContacts.Add(jsonContact);
            }

            if (translate)
            {
                jsonContacts = translateToDisplayName(jsonContacts, "contact");
            }

            responsObject["contacts"] = jsonContacts;

            return responsObject;
        }

        public JArray translateToDisplayName(JArray array, string type)
        {
            var displayNames = crmService.GetAttributeDisplayName(type);

            JArray contactsWithDisplayNames = new JArray();

            foreach (JObject contact in array.Children<JObject>())
            {
                JObject newContact = new JObject();
                foreach (JProperty keyValue in contact.Properties())
                {
                    if (displayNames.ContainsKey(keyValue.Name.ToString().ToLower()))
                    {
                        string displayName;
                        displayNames.TryGetValue(keyValue.Name.ToString().ToLower(), out displayName);

                        if (newContact.Property(displayName) == null)
                        {
                            newContact.Add(displayName, keyValue.Value);
                        }
                        else
                        {
                            newContact.Add(displayName + " 2", keyValue.Value);
                        }

                    }
                }
                contactsWithDisplayNames.Add(newContact);
            }

            return contactsWithDisplayNames;

        }

    }
}