using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Net.Mail;

namespace WordAPIDocAssemblySampleWeb.Controllers
{
    public class DoctorsController : ApiController
    {
        // GET api/<controller>
        public IEnumerable<string> Get()
        {
            return new string[] { "value1", "value2" };
        }

        public class FeedbackRequest
        {
            public int? Rating { get; set; }
            public string Feedback { get; set; }
        }

        public class doctorResponse
        {
            public string id { get; set; }
            public string email { get; set; }
            public string phone { get; set; }
            public string full_name { get; set; }
        }

        [HttpGet()]
        public doctorResponse DoctorsList()
        {
            try
            {
                //const string MailingAddressFrom = "steve.r.frohlich@gmail.com";
                //const string MailingAddressTo = "dev_team@contoso.com";
                //const string SmtpHost = "smtp.contoso.com";
                //const int SmtpPort = 587;
                //const bool SmtpEnableSsl = true;
                //const string SmtpCredentialsUsername = "username";
                //const string SmtpCredentialsPassword = "password";

                var subject = "Sample App feedback, "
                + DateTime.Now.ToString("MMM dd, yyyy, hh:mm tt");

                return new doctorResponse()
                {
                    id = "1",
                    email = "example@example.com",
                    phone = "0761714881",
                    full_name ="Steven Frohlich"
                };

            }
            catch (Exception)
            {
                // Could add some logging functionality here.

                return new doctorResponse()
                {
                    id = "0",
                    email = "failed@example.com",
                    phone = "000000000",
                    full_name = "Failed Failure"
                };
            }
        }

        // GET api/<controller>/5
        public string Get(int id)
        {
            return "value";
        }

        // POST api/<controller>
        public void Post([FromBody]string value)
        {
        }

        // PUT api/<controller>/5
        public void Put(int id, [FromBody]string value)
        {
        }

        // DELETE api/<controller>/5
        public void Delete(int id)
        {
        }
    }
}