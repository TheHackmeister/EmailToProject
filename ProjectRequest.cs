using System;
using System.Net;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Collections.Specialized; // For name/value pairs.
using Newtonsoft.Json.Linq; // For JSON parsing. 


namespace EmailToProject
{
    public class ProjectRequest
    {
        private static String pisces_URL = "http://ubuntu.pcr:8080/";
       
        // At some point, it would make sense to pass in the username and password here. 
        public ProjectRequest()
        {


        }

        // Returns a valid JSON object with error info in it. It can be accessed at the root, as well as in "projects": {}. 
        private JToken communicationError(string text, Exception e) {
            Console.Write(e);
            System.Windows.Forms.MessageBox.Show(text);

            JToken error = JObject.Parse("{\"results\":[{\"title\":\"" + text + "\",\"text\":\"" + text + "\",\"id\":\"-1\",\"project_id\":\"-1\"}]}");            
            return error["results"];
        }

        private WebClient setupClient()
        {
            WebClient client = new WebClient(); // This object gets cleared after reading. Need to recreate.
            client.Headers.Add("accept", "application/json");
            return client;
        }

        public JToken searchProjects(string term)
        {
            WebClient client = setupClient();

            // Only add a star if the thing being searched is a single word. This also avoids emails. 
            string star = term.All(char.IsLetterOrDigit) ? "*" : "";
            try
            {
                string s = client.DownloadString(pisces_URL + "projects" + "?q=" + term + star);
                JObject projects = JObject.Parse(s);
                return projects["projects"];           
            }
            catch (WebException e)
            {
                return communicationError("Error communicating with the server. See the logs for more info.", e);
            }
        }


        public JToken attachEmail(string projectID, string emailAddress, string contactName, string body, string commType, string status)
        {
            // Setup the request.
            WebClient client = setupClient();
            client.Headers[HttpRequestHeader.ContentType] = "application/x-www-form-urlencoded";
            

            // Create our params. 
            NameValueCollection rec = new NameValueCollection();
            rec.Add("communication[project_id]", projectID);
            rec.Add("communication[notes]", body );
// Still need to update this.
            rec.Add("communication[contact_id]", "17"); 
            rec.Add("communication[communication_type_id]", commType); 
            rec.Add("communication[communication_status_id]", status); 
            rec.Add("communication[summary]", "Email"); 

            // Making sure everything went ok when talking with the server.          
            try
            {
                // Run the request and convert it to a string. It comes in as a Byte[]. 
                string s = System.Text.Encoding.ASCII.GetString(client.UploadValues(pisces_URL + "communications", rec));
                JObject result = JObject.Parse(s);
                return result;
            }
            catch (WebException e)
            {
                return communicationError("Error adding the email. See the log for more info.", e);
            }
        }

        public JToken getStatuses()
        {
            WebClient client = setupClient();
            try
            {
                string s = client.DownloadString(pisces_URL + "communication_statuses");
                //System.Windows.Forms.MessageBox.Show(s);
                JObject stats = JObject.Parse(s);
                return stats["communication_statuses"];
            }
            catch (WebException e)
            {
                return communicationError("Error getting statuses. See the logs for more info.", e);
            }
        }

        public JToken getCommType()
        {
            WebClient client = setupClient();
            try
            {
                string s = client.DownloadString(pisces_URL + "communication_types");
                JObject types = JObject.Parse(s);
                return types["communication_types"];
            }
            catch (WebException e)
            {
                return communicationError("Error getting communication types. See the logs for more info.", e);
            }
        }
    }
}
