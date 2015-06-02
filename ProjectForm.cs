using System;
//using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using System.Drawing;
using System.Windows.Forms;
using Newtonsoft.Json.Linq; // For JSON parsing. 

namespace EmailToProject
{
    public class ProjectForm
    {
        private MailItem mail;
        private Button[] projectButtons;
        private Form page;
        private ProjectRequest request;

        // MessageBox.Show(item.SenderEmailAddress + "\n" + item.Subject + "\n" + item.Body + "\n" + item.ReceivedTime);          // display sender email & subject line
        
        public ProjectForm(MailItem mail, ProjectRequest request) {
            this.mail = mail;
            this.request = request;
            createForm();
// Maybe check if the search worked?
            searchEmail();
            page.Show();
        }

        // Creates the form used to search and display projects. 
        private void createForm()
        {
            // Build the form
            page = new Form();
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
            projectButtons = new Button[5];
            for (int i = 0; i < projectButtons.Length;i++ )
            {
                Button button = new Button();
                button.Location = new Point(0, (i + 1) * 23);
                button.Click += attatchToProject;
                button.Visible = false;
                page.Controls.Add(button);
                projectButtons[i] = button;
            }
        }

        private void searchEmail() {
            string json = request.searchProjects(mail.SenderEmailAddress);
            updateButtons(json);
        }

        private void updateButtons(string json)
        {
            JObject projects = JObject.Parse(json);

            int counter = 0;
            foreach (var project in projects["projects"])
            {
                this.projectButtons[counter].Text = (string)project["title"];
                this.projectButtons[counter].Tag = (string)project["id"];
                this.projectButtons[counter].Visible = true;

                // Quit if the array is full. 
                if (counter >= projectButtons.Length)
                {
                    break;
                }
                counter++; 
            }

            // Hides any unused buttons. 
            for (; counter < projectButtons.Length; counter++)
            {
                projectButtons[counter].Visible = false;
            }
        }

        // Attatches an email as a communication to a project. 
        public void attatchToProject(Object sender, EventArgs e)
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
    }
}
