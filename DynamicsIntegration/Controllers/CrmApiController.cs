using System;
using System.Diagnostics;
using System.Web.Http;
using System.Web.Http.Cors;

using DynamicsIntegration.Models;
using DynamicsIntegration.Services;


namespace DynamicsIntegration.Controllers
{

    [RoutePrefix("Dynamics")]
    [EnableCors(origins: "http://172.17.123.104:8888", headers: "*", methods: "*")]
    public class CrmApiController : ApiController
    {
        CrmService crmService;
        CrmApiHelper helper;

        [Route("Lists")]
        public IHttpActionResult GetListsWithAllAttributes([FromUri]DynamicsCredentials credentials, [FromUri] bool translate = false, [FromUri] bool allValues = true)
        {
            try
            {
                crmService = new CrmService(credentials);
                helper = new CrmApiHelper(crmService);
            }
            catch (NullReferenceException)
            {
                return Unauthorized();
            }

            var lists = crmService.getAllLists(allValues);
            var responsObject = helper.getValuesFromLists(lists, translate, allValues);

            return Ok(responsObject);
        }

        [Route("Lists/{listId}/Contacts")]
        public IHttpActionResult GetContactsWithAttributes(string listId, [FromUri]DynamicsCredentials credentials, [FromUri] int preview = 0, [FromUri] bool translate = true, [FromUri] bool allValues = true)
        {
            Debug.WriteLine("contacts anrop " + DateTime.Now.ToString());
            try
            {
                crmService = new CrmService(credentials);
                helper = new CrmApiHelper(crmService); Debug.WriteLine("contacts anrop efter try" + DateTime.Now.ToString());
            }
            catch (NullReferenceException)
            {
                return Unauthorized();
            }

            try
            {
                Debug.WriteLine("contacts anrop innan resp obj " + DateTime.Now.ToString());
                var contacts = crmService.getContactsInList(listId, allValues, preview);
                var responsObject = helper.getValuesFromContacts(contacts, translate, allValues);
                Debug.WriteLine("contacts anrop efter resp obj" + DateTime.Now.ToString());

                return Ok(responsObject);
            }
            catch(NullReferenceException)
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
            catch (NullReferenceException)
            {
                return Unauthorized();
            }

            try
            {
                crmService.changeBulkEmail(contactId);
                return Ok();
            }
            catch (NullReferenceException)
            {
                return BadRequest();
            }
        }

    }
}