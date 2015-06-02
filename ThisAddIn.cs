using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Drawing;
using System.Windows.Forms;
using System.Net;
using System.IO;
using Newtonsoft.Json.Linq; // For JSON parsing. 


namespace EmailToProject
{
    public partial class EmailToProject
    {
        private static String pisces_URL = "http://ubuntu.pcr:8080/projects";
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
// Probably get Auth token here. 
            Application.ItemContextMenuDisplay += ApplicationItemContextMenuDisplay;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
// Give up Auth token here. 
        }

        // Outlook menu item
        // This triggers whenever a mailitem is rightclicked, and gets a "selection" object passed which contains all selected items
        void ApplicationItemContextMenuDisplay(CommandBar commandBar, Selection selection)
        {
            var cb = commandBar.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, missing, true) as Office.CommandBarButton;
            if (cb == null) return;
            cb.Visible = true;
            //cb.Picture = ImageConverter.ImageToPictureDisp(Properties.Resources.Desktop.ToBitmap());    // some icon stored in the resources file
            cb.Style = MsoButtonStyle.msoButtonIconAndCaption;                                          // set style to text AND icon
            cb.Click += new _CommandBarButtonEvents_ClickEventHandler(AsterixHook);                     // link click event

            // single MailItem item selection only, NOT 0 based
            if (selection.Count == 1 && selection[1] is Outlook.MailItem)
            {
                var item = (MailItem)selection[1];                          // retrieve the selected item
                string subject = item.Subject;
                if (subject.Length > 25) subject = subject.Substring(0, 25);// limit max length of the caption
                cb.Caption = "Import into Projects";                         // set caption
                cb.Enabled = true;                                          // user selected a single mail item, enable the menu
                cb.Parameter = item.EntryID;                                // this will pass the selected item's identification down when clicked
            }
            else
            {
                cb.Caption = "Invalid selection";
                cb.Enabled = false;
            }

        }


        private void AsterixHook(CommandBarButton control, ref bool canceldefault)
        {
            string entryid = control.Parameter;                                     // the outlook entry id clicked by the user
            var item = (MailItem)this.Application.Session.GetItemFromID(entryid);   // the actual item


           

            WebClient client = new WebClient();
            client.Headers.Add("accept", "application/json");


            // Sends the request
            Stream data = client.OpenRead(pisces_URL + "?q=" + item.SenderEmailAddress);
  
            // Gets the results. Is there a thread lock here?
            // Yes :(. 
            StreamReader reader = new StreamReader(data);
            // Need to check if we errored out here. 

            string s = reader.ReadToEnd();

            // Clean it up. 
            data.Close();
            reader.Close();

            //dynamic projects = Newtonsoft.Json.JsonConvert.DeserializeObject(s);
            // MessageBox.Show(item.SenderEmailAddress + "\n" + item.Subject + "\n" + item.Body + "\n" + item.ReceivedTime);          // display sender email & subject line
            
            // Build the form
            Form page = new Form();
            page.Text = "Add email to project";
            page.Width = 380;
            page.Height = 155;
            
            // Search box
            TextBox tb = new TextBox();
            tb.Width = 250;
            page.Controls.Add(tb);

            // Search button
            Button search = new Button();
            search.Text = "Search Projects";
            search.Location = new Point(260, 0);
            search.Width = 100;
            page.Controls.Add(search);

            // Search results 
            Button[] projectButtons = new Button[5];


            JObject projects = JObject.Parse(s);
                 
            

            int numButtons = 0; // Counter. 
            foreach (var project in projects["projects"])
            {

                Button button = new Button();
                button.Text = (string)project["title"];
                button.Tag = (string)project["id"];
                button.Location = new Point(0, (numButtons + 1) * 23);
                button.Click += attatchToProject;
                projectButtons[numButtons] = button;

                // Quit if the array is full. 
                if (numButtons >= projectButtons.Length)
                {
                    break;
                }
                numButtons++; 
            }
   /*         */
            //MessageBox.Show(s);

            // Creating the interface. 
            
            

            for(int i = 0; i < numButtons; i ++) {
                page.Controls.Add(projectButtons[i]);
                // MessageBox.Show("Adding button: " + i);
                
            }

            page.Show();

            


            // Create button, add to page.Show page.
        }

        // Attatches an email as a communication to a project. 
        private void attatchToProject(Object sender, EventArgs e)
        {
            Button button = (Button)sender;

            // Get project ID. 
            // Get email address. 
            // Get contact name from email.
            // Get body.
            // Make put request. (Let rails handle creating a contact, if needed).
            // Show outcome. 
            
            button.Text = (string)button.Tag;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}







