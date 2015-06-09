using System;
//using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using System.Drawing;
using System.Windows.Forms;
using Newtonsoft.Json.Linq; // For JSON parsing. 
//using System.Threading.Tasks; // For requests. 
using System.Collections.Generic;

namespace EmailToProject
{
    public class ProjectForm
    {
        private static Dictionary<string, string> statuses;
        private static string commType;
        private MailItem mail;
        private Button[] projectButtons;
        private Form page;
        private ProjectRequest request;
        private TextBox tb;
        private ComboBox cmbox;
        private string defaultStatus = "Returned";

        public ProjectForm(MailItem mail, ProjectRequest request) {
            this.mail = mail;
            this.request = request;
            prepForm();
            createForm();
            searchEmail();
            page.Show();
        }

        public void prepForm()
        {
            setupStatuses();
            setupCommType();
        }

        // Creates the form used to search and display projects. 
        private void createForm()
        {
            // Build the form
            page = new Form();
            page.Text = "Add email to project";
            page.Width = 380;
            page.Height = 180;

            // Search box
            tb = new TextBox();
            tb.Width = 250;
            tb.KeyDown += enterFromSearch;
            page.Controls.Add(tb);

            // Search button
            Button search = new Button();
            search.Text = "Search Projects";
            search.Location = new Point(260, 0);
            search.Width = 100;
            search.Click += searchProjects;
            page.Controls.Add(search);

            // Status Type
            cmbox = new ComboBox();
            cmbox.DropDownStyle = ComboBoxStyle.DropDownList; // Make it a dropdown. 
            cmbox.Location = new Point(260, 24);
            cmbox.Width = 100;
            cmbox.Items.AddRange(System.Linq.Enumerable.ToArray(statuses.Keys));
            cmbox.SelectedItem = "Returned"; // I have to take only the names from the dictionary because otherwise I can't select the item I want.
            page.Controls.Add(cmbox);


            // Search results 
            projectButtons = new Button[5];
            for (int i = 0; i < projectButtons.Length;i++ )
            {
                Button button = new Button();
                button.Width = 250;
                button.Location = new Point(0, (i + 1) * 23);
                button.Click += attatchToProject;
                button.Visible = false;
                page.Controls.Add(button);
                projectButtons[i] = button;
            }
        }

        // Makes sure statuses is initilized and not in error. If it is, initilize it. 
        private void setupStatuses()
        {
            if (statuses == null || statuses.ContainsKey("-1"))
            {
                statuses = new Dictionary<string, string>();
                JToken stats = request.getStatuses();
                foreach (var status in stats)
                {
                    statuses.Add((string)status.SelectToken("text"), (string)status.SelectToken("id"));
                }
            }
        }
        
        // Makes sure the communication type is able to be set to Email.
        private void setupCommType()
        {
            if (commType == null || commType == "-1")
            {
                JToken jCommType = request.getCommType();
                foreach (var cType in jCommType) 
                {
                    if((string)cType.SelectToken("text") == "Email") {
                        commType = (string)cType.SelectToken("id");
                        break;
                    }
                }

// Error out here?
            }
        }

        private void searchEmail() {
            JToken json = request.searchProjects(mail.SenderEmailAddress);
            updateButtons(json);
        }

        private void enterFromSearch(Object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                searchProjects(sender, e);
            }
        }

        private void searchProjects(Object sender, EventArgs e)
        {
            JToken json = request.searchProjects(tb.Text);
            updateButtons(json);
        }

        private void updateButtons(JToken projects)
        {
            int counter = 0;
            foreach (var project in projects)
            {
                this.projectButtons[counter].Text = (string)project.SelectToken("title");
                this.projectButtons[counter].Tag = (string)project.SelectToken("id");
                this.projectButtons[counter].Visible = true;

                counter++;
                // Quit if the array is full. 
                if (counter >= projectButtons.Length)
                {
                    break;
                }
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
            string id = (string)button.Tag;
            string email = mail.SenderEmailAddress;
            string contact = mail.SenderName;
            string body = mail.Subject + "\n" + mail.Body;
           
            JToken status = request.attachEmail(id, email, contact, body, commType, statuses[(string)cmbox.SelectedItem]);
            page.Hide();


            
            if ((string)status.SelectToken("project_id") == id)
            {
                MessageBox.Show("Email added successfully.");
                // If there is an error, the error handler is responsible for displaying it. 
            }
        
        }
    }
}
