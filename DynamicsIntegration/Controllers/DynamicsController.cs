using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using RestSharp;
using DynamicsIntegration.Models;

namespace DynamicsIntegration.Controllers
{
    [RoutePrefix("Dynamics")]
    public class DynamicsController : ApiController
    {
        [Route("ListSet")]
        public IHttpActionResult GetListSets()
        {
            // Connect to Dynamics
            // Get views
            Item listSets = GetListSetsFromDynamics();
            //var modified = listSets.Select(v => new listSets
            //{
            //    ViewId = v.viewId,
            //    Name = v.name
            //};)

            return Ok("tendeee");
        }

        public class View
        {
            Guid ViewId { get; set; }

            string Name { get; set; }
        }

        private dynamic GetListSetsFromDynamics()
        {

            var client = new RestClient("http://brottsplatskartan.se");
            // client.Authenticator = new HttpBasicAuthenticator(username, password);
           
            var request = new RestRequest("api.php?action=getEvents&period=1440", Method.GET);

            //// easily add HTTP Headers
            //request.AddHeader("header", "value");

            // execute the request
            List<Item> response = client.Execute<List<Item>>(request).Data;
            foreach(Item item in response){
                System.Console.WriteLine(item.Title);
            }

            Item content = response.FirstOrDefault<Item>();
           // var content = response.First<Item> // raw content as string
            
            return content;
        }

        /*[Route("Views/{view}")]
        public dynamic GetData(string view)
        {

        }*/
    }
}