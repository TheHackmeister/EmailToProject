using System;
using System.Net;
using System.IO;
using System.Threading.Tasks;

namespace EmailToProject
{
    public class ProjectRequest
    {
        private static String pisces_URL = "http://ubuntu.pcr:8080/projects";
       
        // At some point, it would make sense to pass in the username and password here. 
        public ProjectRequest()
        {


        }

        public string searchProjects(string term)
        {
            WebClient client = new WebClient(); // This object gets cleared after reading. Need to recreate.
            client.Headers.Add("accept", "application/json");
            string s = client.DownloadString(pisces_URL + "?q=" + term) + "*"; // I add the star to be able to shorten scintific terms. For other things, it shouldn't have a large impact. 
            return s;
        }


        public string attachEmail(string projectID, string emailAddress, string contactName, string body)
        {
            WebClient client = new WebClient(); // This object gets cleared after reading. Need to recreate.
            client.Headers[HttpRequestHeader.ContentType] = "application/x-www-form-urlencoded";
            string s = client.UploadString(pisces_URL, "id=" + projectID);
            return s;
        }
    }
}
