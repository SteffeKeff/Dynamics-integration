using Microsoft.Xrm.Sdk;
using Newtonsoft.Json.Linq;
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

        public JArray getValuesFromLists(EntityCollection lists, bool translate, bool allAttributes)
        {
            var jsonLists = new JArray();
            
            foreach (var list in lists.Entities)
            {
                var jsonList = new JObject();

                foreach (var prop in list.GetType().GetProperties())
                {
                    if (prop.Name != "Item")
                    {
                        var val = prop.GetValue(list, null);

                        if (allAttributes | val != null)
                            jsonList[prop.Name] = val == null ? null : val.ToString();
                    }
                }

                jsonLists.Add(jsonList);
            }

            if (translate)
            {
                jsonLists = translateToDisplayName(jsonLists, "list");
            }

            return jsonLists;
        }

        public JArray getValuesFromContacts(EntityCollection contacts, bool translate, bool allAttributes)
        {
            var jsonContacts = new JArray();

            foreach (var contact in contacts.Entities)
            {
                var jsonContact = new JObject();

                foreach (var prop in contact.GetType().GetProperties())
                {
                    if (prop.Name != "Item")
                    {
                        var val = prop.GetValue(contact, null);

                        if (allAttributes | val != null)
                            jsonContact[prop.Name] = val == null ? null : val.ToString();
                    }
                }

                jsonContacts.Add(jsonContact);
            }

            if (translate)
            {
                jsonContacts = translateToDisplayName(jsonContacts, "contact");
            }

            return jsonContacts;
        }

        public JArray translateToDisplayName(JArray array, string type)
        {
            var displayNames = crmService.GetAttributeDisplayName(type);
            var contactsWithDisplayNames = new JArray();

            foreach (var contact in array.Children<JObject>())
            {
                var newContact = new JObject();
                foreach (var keyValue in contact.Properties())
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