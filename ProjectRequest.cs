using System;
using System.Net;
using System.IO;

namespace EmailToProject
{
    public class ProjectRequest
    {
        private static String pisces_URL = "http://ubuntu.pcr:8080/projects";
        private WebClient client;


        // At some point, it would make sense to pass in the username and password here. 
        public ProjectRequest()
        {
            // I actually don't know if this should stick around or if it's a one and done thing. 
            client = new WebClient();
            client.Headers.Add("accept", "application/json");

        }

        public string searchProjects(string term) {
            // Sends the request
            Stream data = client.OpenRead(pisces_URL + "?q=" + term);
// Gets the results. Is there a thread lock here?
// Yes :(. 
            StreamReader reader = new StreamReader(data);
// Need to check if we errored out here. 

            string s = reader.ReadToEnd();
            
            // Clean it up. 
            data.Close();
            reader.Close();
            return s;
        }
    }
}
