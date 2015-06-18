using System;
using System.Net;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Collections.Specialized; // For name/value pairs.

using Newtonsoft.Json.Linq; // For JSON parsing. 
//using Microsoft.Office.Interop.Outlook; // For StorageItem.
using Outlook = Microsoft.Office.Interop.Outlook; //For StorageItem.
using System.Windows; // For StorageItem?

namespace EmailToProject
{
    public class ProjectRequest
    {
        private static String pisces_URL = "http://ubuntu.pcr:8080/";
        private Outlook.StorageItem storage;
        private string email;
        private string token;
        private string error;
        private Boolean authStatus;


        // At some point, it would make sense to pass in the username and password here. 
        public ProjectRequest()
        {
            error = "";
            authStatus = false;
        }

        public string getError()
        {
            string oldError = error != "" ? error : "";
            error = "";
            return oldError;
        }

        // Returns a valid JSON object with error info in it. It can be accessed at the root, as well as in "projects": {}. 
        private JToken communicationError(string text, Exception e)
        {
            error = text;

            JToken tokenError = JObject.Parse("{\"results\":[{\"title\":\"" + text + "\",\"text\":\"" + text + "\",\"id\":\"-1\",\"project_id\":\"-1\"}]}");
            return tokenError["results"];
        }

        private void loadAuth() 
        {
            try
            {
                //Outlook.Folder oInbox = Outlook.Application
                var outlookApp = new Outlook.Application();

                Outlook.StorageItem storage = outlookApp.Session.GetDefaultFolder(
                    Outlook.OlDefaultFolders.olFolderCalendar).GetStorage(
                    "PiscesProjects",
                    Outlook.OlStorageIdentifierType.olIdentifyByMessageClass);
                Outlook.PropertyAccessor pa = storage.PropertyAccessor;
                
                // PropertyAccessor will return a byte array for this property

                email = pa.GetProperty("http://schemas.microsoft.com/mapi/string/{FFF40745-D92F-4C11-9E14-92701F001EB3}/email");
                token = pa.GetProperty("http://schemas.microsoft.com/mapi/string/{FFF40745-D92F-4C11-9E14-92701F001EB3}/token");
            }
            catch
            {
                return;
            }
        }
        
        private void storeAuth(string newEmail, string newToken)
        {
            //Outlook.Folder oInbox = Outlook.Application
            var outlookApp = new Outlook.Application();

            Outlook.StorageItem storage = outlookApp.Session.GetDefaultFolder(
                Outlook.OlDefaultFolders.olFolderCalendar).GetStorage(
                "PiscesProjects",
                Outlook.OlStorageIdentifierType.olIdentifyByMessageClass);
            Outlook.PropertyAccessor pa = storage.PropertyAccessor;

            pa.SetProperty("http://schemas.microsoft.com/mapi/string/{FFF40745-D92F-4C11-9E14-92701F001EB3}/email", newEmail);
            pa.SetProperty("http://schemas.microsoft.com/mapi/string/{FFF40745-D92F-4C11-9E14-92701F001EB3}/token", newToken);
            storage.Save();
        }

        private string getToken()
        {
            if (token == null) loadAuth();
            return token;
        }

        private string getEmail()
        {
            if (email == null) loadAuth();
            return email;
        }

        private string getAuth()
        {
            return "token=" + getToken() + "&email=" + getEmail();
        }

        private WebClient setupClient()
        {
            WebClient client = new WebClient(); // This object gets cleared after reading. Need to recreate.
            client.Headers.Add("accept", "application/json");
            return client;
        }

        private JToken makeGetRequest(string path, string options, string requestError, string tickler = "")
        {

            authReady(); // This will load credentials. If they don't work, this will fail in the try block, erroring out.
            if (error != "") return true;

            WebClient client = setupClient();
            string fullOptions = options != "" ? "?" + options + "&" : "?";
            try
            {
                string s = client.DownloadString(pisces_URL + path + fullOptions + getAuth());
                JObject obj = JObject.Parse(s);
                tickler = tickler == "" ? path : tickler;
                return obj[tickler];
            }
            catch (WebException e)
            {
// I should do a better job of catching what the error is here. 
                authStatus = false;
                return communicationError(requestError, e);
            }
        }

        private Boolean authReady()
        {
            if (authStatus == false)
            {
                loadAuth();
                if (email == "" || token == "")
                {
                    error = "Invalid credentials"; // "No credentials saved. Please login.";
                    return false;
                }
                authStatus = true;
                JToken result = makeGetRequest("token/show", "", "Invalid credentials", "user"); // Matching for "Invalid credentials" is used else where. Just be aware. 
                if (error != "" ) return false;
            }
            return true;
        }

        public void forgetAuth()
        {
            storeAuth("", "");
            loadAuth();
        }

        public JToken getStatuses()
        {
            return makeGetRequest("communication_statuses", "", "Error getting statuses. See the logs for more info.");
        }

        public JToken getCommType()
        {
            return makeGetRequest("communication_types", "", "Error getting communication types. See the logs for more info.");
        }

        public JToken searchProjects(string term)
        {
            string star = term.All(char.IsLetterOrDigit) ? "*" : "";
            return makeGetRequest("projects", "q=" + term + star, "Error communicating with the server. See the logs for more info.");
        }

        public Boolean getToken(string email, string password)
        {
            WebClient client = setupClient();
            try
            {
                string s = client.DownloadString(pisces_URL + "token/show" + "?email=" + email + "&password=" + password);
                
                JObject tok = JObject.Parse(s);
                JToken auth = tok["users"];
                
                if (auth.ToString() == "") return true; // An error!
                storeAuth(email, auth["token"].ToString());
                return false;
            }
            catch (WebException e)
            {
                communicationError("Error retriving your authorization token.", e);
                return true; // An error. 
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
            rec.Add("communication[notes]", body);
            rec.Add("contact_email", emailAddress);
            rec.Add("contact_name", contactName); 
            rec.Add("communication[communication_type_id]", commType); 
            rec.Add("communication[communication_status_id]", status); 
            rec.Add("communication[summary]", "Email");
            rec.Add("email", getEmail());
            rec.Add("token", getToken());

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
    }
}
