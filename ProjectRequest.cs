using System;
using System.Net;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Collections.Specialized; // For name/value pairs.

namespace EmailToProject
{
    public class ProjectRequest
    {
        private static String pisces_URL = "http://ubuntu.pcr:8080/";
       
        // At some point, it would make sense to pass in the username and password here. 
        public ProjectRequest()
        {


        }

        public string searchProjects(string term)
        {
            WebClient client = new WebClient(); // This object gets cleared after reading. Need to recreate.
            client.Headers.Add("accept", "application/json");
            
            string s = client.DownloadString(pisces_URL + "projects" + "?q=" + term + addStar(term) ); 
            return s;
        }

        // Only add a star if the thing being searched is a single word. This also avoids emails. 
        private string addStar(string term)
        {
            if (term.All(char.IsLetterOrDigit))
            {
                return "*";
            }
            return "";
        }

        public string attachEmail(string projectID, string emailAddress, string contactName, string body)
        {
            // Setup the request.
            WebClient client = new WebClient(); // This object gets cleared after reading. Need to recreate.
            client.Headers[HttpRequestHeader.ContentType] = "application/x-www-form-urlencoded";
            client.Headers.Add("accept", "application/json");

            // Create our params. 
            NameValueCollection rec = new NameValueCollection();
            rec.Add("communication[project_id]", projectID);
            rec.Add("communication[notes]", body );
            rec.Add("communication[contact_id]", "17");
// Need some way of ensuring email is of type 2. 
            rec.Add("communication[communication_type_id]", "2"); 
            rec.Add("communication[communication_status_id]", "1"); // I don't know what 1 is. 
            rec.Add("communication[summary]", "Email"); 

            // Run the request and convert it to a string. It comes in as a Byte[]. 
            string s = System.Text.Encoding.ASCII.GetString(client.UploadValues(pisces_URL + "communications", rec));
            
            return s;
        }
    }
}
