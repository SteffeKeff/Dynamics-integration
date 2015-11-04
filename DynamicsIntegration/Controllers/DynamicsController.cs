using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using RestSharp;
using System.Diagnostics;
using DynamicsIntegration.Models;
using Microsoft.Xrm.Sdk;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

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

            //var listSets = GetListSetsFromDynamics();


            //var modified = listSets.Select(v => new listSets
            //{
            //    ViewId = v.viewId,
            //    Name = v.name
            //};)

            AuthorityController ac = new AuthorityController();
            EntityCollection results = ac.getSomething();

            var obj = new JObject();

            //obj["One"] = "Value One";
            //obj["Two"] = "Value Two";
            //obj["Three"] = "Value Three";

            var array = new JArray();

            foreach (List lista in results.Entities)
            {
                var list = new JObject();
                list["name"] = lista.ListName;
                list["listid"] = lista.MemberCount;
                list["membercount"] = lista.ListId;
                list["modifiedon"] = lista.ModifiedOn;
                /*
                Console.WriteLine("Title: {0}", lista.ListName);
                        Console.WriteLine("count: {0}", lista.MemberCount);
                        Console.WriteLine("Id: {0}", lista.ListId);
                        Console.WriteLine("ModifiedOn: {0}", lista.ModifiedOn);
                        */

                array.Add(list);
            }

            obj["marketinglist"] = array;

            return Ok(obj);

            //return Ok("hej");
        }

        public class View
        {
            Guid ViewId { get; set; }

            string Name { get; set; }
        }

        private dynamic GetListSetsFromDynamics()
        {

            var client = new RestClient("https://ungapped.crm4.dynamics.com");
            // client.Authenticator = new HttpBasicAuthenticator(username, password);
           
            var request = new RestRequest("XRMServices/2011/OrganizationData.svc/ListSet", Method.GET);

            //// easily add HTTP Headers
            //request.AddHeader("header", "value");

            // execute the request
            var response = client.Execute(request).Content;
            Debug.WriteLine(response);

            // var content = response.First<Item> // raw content as string

            return response;
        }

        /*[Route("Views/{view}")]
        public dynamic GetData(string view)
        {

        }*/
            }
        }