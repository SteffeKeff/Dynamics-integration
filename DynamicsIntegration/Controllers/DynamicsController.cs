using DynamicsIntegration.Models;
using Microsoft.Xrm.Sdk;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Web.Http;
using System.Collections;
using System.Web.Http.Cors;
using System.Reflection;

namespace DynamicsIntegration.Controllers
{
    [RoutePrefix("dynamics")]
    [EnableCors(origins: "http://172.17.123.104:8888", headers: "*", methods: "*")]
    public class DynamicsController : ApiController
    {
        AuthorityController authorityController;

        [Route("displaynames/{entitySchemaName}")]
        public IHttpActionResult GetDisplayName(string entitySchemaName, [FromUri]DynamicsCredentials credentials)
        {    

            try
            {
                authorityController = new AuthorityController(credentials);
            }
            catch(Exception)
            {
                return Unauthorized();
            }
            var localNames = authorityController.GetAttributeDisplayName(entitySchemaName);

            return Ok(localNames);
        }

        [Route("lists")]
        public IHttpActionResult GetLists([FromUri]DynamicsCredentials credentials)
        {

            try
            {
                authorityController = new AuthorityController(credentials);
            }
            catch (Exception)
            {
                return Unauthorized();
            }

            EntityCollection lists = authorityController.getAllLists(false);
            var responsObject = getValuesFromLists(lists, false);

            return Ok(responsObject);
        }

        [Route("lists2")]
        public IHttpActionResult GetListsWithAllAttributes([FromUri]DynamicsCredentials credentials)
        {

            try
            {
                authorityController = new AuthorityController(credentials);
            }
            catch (Exception)
            {
                return Unauthorized();
            }

            EntityCollection lists = authorityController.getAllLists(true);
            var responsObject = getValuesFromLists(lists, true);

            return Ok(responsObject);
        }
        
        /// /Lists/id/Contacts?Domain=X&UserName=Y&Password=Z
        [Route("Lists/{listId}/Contacts")]
        public IHttpActionResult GetContacts(string listId, [FromUri]DynamicsCredentials credentials, [FromUri] int preview = 0)
        {

            try
            {
                authorityController = new AuthorityController(credentials);
            }
            catch (Exception)
            {
                return Unauthorized();
            }

            try
            {
                ArrayList contacts = authorityController.getContactsInList(listId, false, preview);
                var responsObject = getValuesFromContacts(contacts, false);
                return Ok(responsObject);
            }
            catch(Exception)
            {
                return BadRequest();
            }
            
        }

        [Route("lists2/{listId}/contacts")]
        public IHttpActionResult GetContactsWithAttributes(string listId, [FromUri]DynamicsCredentials credentials,[FromUri] int preview = 0)
        {

            try
            {
                authorityController = new AuthorityController(credentials); 
            }
            catch (Exception)
            {
                return Unauthorized();
            }

            try
            {
                ArrayList contacts = authorityController.getContactsInList(listId, true, preview);
                var responsObject = getValuesFromContacts(contacts, true);

                return Ok(responsObject);
            }
            catch(Exception)
            {
                return BadRequest();
            }
        }

        private JObject getValuesFromLists(EntityCollection lists, bool allAttributes)
        {
            var responsObject = new JObject();
            var jsonLists = new JArray();

            foreach (List list in lists.Entities)
            {
                var jsonList = new JObject();

                if(allAttributes)
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
            jsonLists = translateToDisplayName(jsonLists, "list");

            responsObject["lists"] = jsonLists;

            return responsObject;
        }

        private JObject getValuesFromContacts(ArrayList contacts, bool allAttributes)
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

            jsonContacts = translateToDisplayName(jsonContacts, "contact");

            responsObject["contacts"] = jsonContacts;

            return responsObject;
        }

        private JArray translateToDisplayName(JArray array, string type)
        {
            var displayNames = authorityController.GetAttributeDisplayName(type);

            JArray arrayWithDisplayNames = new JArray();

            foreach(JObject contact in array.Children<JObject>())
            {
                JObject newContact = new JObject();
                foreach(JProperty keyValue in contact.Properties())
                {
                    if(displayNames.ContainsKey(keyValue.Name.ToString().ToLower()))
                    {
                        string displayName;
                        displayNames.TryGetValue(keyValue.Name.ToString().ToLower(), out displayName);

                        try
                        {
                            newContact.Add(displayName, keyValue.Value);
                        }
                        catch(Exception)
                        {
                            
                        }
                        
                    }
                }
                arrayWithDisplayNames.Add(newContact);
            }

            return arrayWithDisplayNames;

        }

    }
}