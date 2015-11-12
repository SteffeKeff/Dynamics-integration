using System;
using System.Web.Http;
using System.Web.Http.Cors;
using System.ServiceModel;
using System.ServiceModel.Security;

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

        [Route("Validate")]
        [HttpPost]
        public IHttpActionResult ValidateCredentials([FromUri]DynamicsCredentials credentials)
        {
            try
            {
                crmService = new CrmService(credentials);

                return Ok(crmService.getOrganizations());
            }
            catch (Exception ex) when (ex is MessageSecurityException || ex is ArgumentNullException)
            {
                return Unauthorized();
            }
            catch(Exception ex)
            {
                return InternalServerError(ex);
            }
        }

        [Route("Organizations/{orgName}/MarketLists")]
        [HttpPost]
        public IHttpActionResult GetListsWithAllAttributes(string orgName, [FromUri]DynamicsCredentials credentials, [FromUri] bool translate = false, [FromUri] bool allValues = true)
        {
            try
            {
                crmService = new CrmService(credentials, orgName);
                helper = new CrmApiHelper(crmService);

                var lists = crmService.getAllLists(allValues);
                var responsObject = helper.getValuesFromLists(lists, translate, allValues);

                return Ok(responsObject);
            }
            catch (Exception ex) when (ex is MessageSecurityException || ex is ArgumentNullException)
            {
                return Unauthorized();
            }
            catch(NullReferenceException)
            {
                return BadRequest();
            }
        }

        [Route("Organizations/{orgName}/MarketLists/{listId}/Contacts")]
        [HttpPost]
        public IHttpActionResult GetContactsWithAttributes(string orgName, string listId, [FromUri]DynamicsCredentials credentials, [FromUri] int top = 0, [FromUri] bool translate = true, [FromUri] bool allValues = true)
        {
            try
            {
                crmService = new CrmService(credentials, orgName);
                helper = new CrmApiHelper(crmService);

                var contacts = crmService.getContactsInList(listId, allValues, top);
                var responsObject = helper.getValuesFromContacts(contacts, translate, allValues);

                return Ok(responsObject);
            }
            catch (Exception ex) when (ex is MessageSecurityException || ex is ArgumentNullException)
            {
                return Unauthorized();
            }
            catch (Exception ex) when (ex is NullReferenceException || ex is FormatException)
            {
                return BadRequest();
            }
        }

        [Route("Organizations/{orgName}/Contacts/{contactId}/Donotbulkemail")]
        [HttpPut]
        public IHttpActionResult UpdateBulkEmailForContact(string orgName, string contactId, [FromUri]DynamicsCredentials credentials)
        {
            try
            {
                crmService = new CrmService(credentials, orgName);

                crmService.changeBulkEmail(contactId);

                return Ok();
            }
            catch (Exception ex) when (ex is MessageSecurityException || ex is ArgumentNullException)
            {
                return Unauthorized();
            }
            catch (Exception ex) when (ex is NullReferenceException || ex is FormatException || ex is FaultException)
            {
                return BadRequest();
            }
        }
    }
}