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

namespace DynamicsIntegration.Controllers
{
    [RoutePrefix("dynamics")]
    public class DynamicsController : ApiController
    {
        AuthorityController authorityController = new AuthorityController();

        [Route("lists")]
        public IHttpActionResult GetLists()
        {
            EntityCollection lists = new EntityCollection();
            try
            {
                lists = authorityController.getAllLists();
            }
            catch(Exception)
            {
                return Unauthorized();
            }
            var responsObject = new JObject();
            var jsonLists = new JArray();

            foreach (List list in lists.Entities)
            {
                var jsonList = new JObject();
                jsonList["name"] = list.ListName;
                jsonList["listid"] = list.ListId;
                jsonList["membercount"] = list.MemberCount;
                jsonList["modifiedon"] = list.ModifiedOn;

                jsonLists.Add(jsonList);
            }

            responsObject["lists"] = jsonLists;

            return Ok(responsObject);
        }

        [Route("lists/{id}")]
        public IHttpActionResult GetSingleList(string id)
        {
            ArrayList contacts = new ArrayList();

            try
            {
                contacts = authorityController.getContactsInList(id);
            }
            catch(Exception)
            {
                return Unauthorized();
            }
            
            var responsObject = new JObject();
            var jsonContacts = new JArray();

            foreach (Contact contact in contacts)
            {
                var jsonContact = new JObject();
                jsonContact["firstname"] = contact.FirstName;
                jsonContact["lastname"] = contact.LastName;
                jsonContact["email"] = contact.EMailAddress1;

                jsonContacts.Add(jsonContact);
            }

            if(contacts.Capacity == 0)
            {
                return NotFound();
            }

            responsObject["contacts"] = jsonContacts;

            return Ok(responsObject);
        }

        [HttpPost]
        [Route("login")]
        public IHttpActionResult Login([FromBody] JObject credentials)
        {

            try
            {
                AuthorityController._domain = credentials.GetValue("domain").ToString();
                AuthorityController._userName = credentials.GetValue("username").ToString();
                AuthorityController._password = credentials.GetValue("password").ToString();

                bool loggedIn = authorityController.login();

                if (loggedIn)
                {
                    return Ok();
                }
                return Unauthorized();
            }
            catch(NullReferenceException)
            {
                return BadRequest();
            }
        }

    }
}