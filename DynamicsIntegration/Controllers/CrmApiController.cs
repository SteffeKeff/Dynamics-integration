using DynamicsIntegration.Models;
using Microsoft.Xrm.Sdk;
using System;
using System.Web.Http;
using System.Collections;
using System.Web.Http.Cors;
using DynamicsIntegration.Services;

namespace DynamicsIntegration.Controllers
{
    [RoutePrefix("Dynamics")]
    [EnableCors(origins: "http://172.17.123.104:8888", headers: "*", methods: "*")]
    public class CrmApiController : ApiController
    {
        CrmService crmService;
        CrmApiHelper helper;

        [Route("Displaynames/{entitySchemaName}")]
        public IHttpActionResult GetDisplayName(string entitySchemaName, [FromUri]DynamicsCredentials credentials)
        {    

            try
            {
                crmService = new CrmService(credentials);
            }
            catch(Exception)
            {
                return Unauthorized();
            }
            var localNames = crmService.GetAttributeDisplayName(entitySchemaName);

            return Ok(localNames);
        }

        [Route("Lists")]
        public IHttpActionResult GetLists([FromUri]DynamicsCredentials credentials, [FromUri] bool translate = true)
        {

            try
            {
                crmService = new CrmService(credentials);
                helper = new CrmApiHelper(crmService);
            }
            catch (Exception)
            {
                return Unauthorized();
            }

            EntityCollection lists = crmService.getAllLists(false);
            var responsObject = helper.getValuesFromLists(lists, translate, false);

            return Ok(responsObject);
        }

        [Route("Lists2")]
        public IHttpActionResult GetListsWithAllAttributes([FromUri]DynamicsCredentials credentials, [FromUri] bool translate = true)
        {

            try
            {
                crmService = new CrmService(credentials);
                helper = new CrmApiHelper(crmService);
            }
            catch (Exception)
            {
                return Unauthorized();
            }

            EntityCollection lists = crmService.getAllLists(true);
            var responsObject = helper.getValuesFromLists(lists, translate, true);

            return Ok(responsObject);
        }
        
        [Route("Lists/{listId}/Contacts")]
        public IHttpActionResult GetContacts(string listId, [FromUri]DynamicsCredentials credentials, [FromUri] int preview = 0, [FromUri] bool translate = true)
        {

            try
            {
                crmService = new CrmService(credentials);
                helper = new CrmApiHelper(crmService);
            }
            catch (Exception)
            {
                return Unauthorized();
            }

            try
            {
                ArrayList contacts = crmService.getContactsInList(listId, false, preview);
                var responsObject = helper.getValuesFromContacts(contacts, translate, false);

                return Ok(responsObject);
            }
            catch(Exception)
            {
                return BadRequest();
            }
        }

        [Route("Lists2/{listId}/Contacts")]
        public IHttpActionResult GetContactsWithAttributes(string listId, [FromUri]DynamicsCredentials credentials, [FromUri] int preview = 0, [FromUri] bool translate = true)
        {

            try
            {
                crmService = new CrmService(credentials);
                helper = new CrmApiHelper(crmService);
            }
            catch (Exception)
            {
                return Unauthorized();
            }

            try
            {
                ArrayList contacts = crmService.getContactsInList(listId, true, preview);
                var responsObject = helper.getValuesFromContacts(contacts, translate, true);

                return Ok(responsObject);
            }
            catch(Exception)
            {
                return BadRequest();
            }
        }

        [Route("Contacts/{contactId}/donotbulkemail")]
        [HttpPut]
        public IHttpActionResult UpdateBulkEmailForContact(string contactId, [FromUri]DynamicsCredentials credentials)
        {

            try
            {
                crmService = new CrmService(credentials);
            }
            catch (Exception)
            {
                return Unauthorized();
            }

            try
            {
                string result = crmService.changeBulkEmail(contactId);

                return Ok(result);
            }
            catch (Exception)
            {
                return BadRequest();
            }
        }

    }
}