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
        private Form login;
        private TextBox userTB;
        private TextBox passTB;
        private string defaultStatus = "Returned";
        

        public ProjectForm(MailItem mail, ProjectRequest request) {
            this.mail = mail;
            this.request = request;
            if(prepForm() == true) return; // If there is an error, it will be displayed. Stop proccessing. 
            createForm();
            if(searchEmail() == true) return;
            page.Show();
        }

        public Boolean prepForm()
        {
            //if (setupAuth() == true) return true;
            if (setupStatuses() == true) return true;
            if (setupCommType() == true) return true;
            return false;
        }

        private Boolean checkRequestError()
        {
            string errorMessage = request.getError();

            if (errorMessage == "Invalid credentials")
            {
                getAuth();
                return true;
            }

            if (errorMessage != "")
            {
                MessageBox.Show(errorMessage);
                return true;
            }

            return false;
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

            // Logout button
            Button logout = new Button();
            logout.Text = "Logout";
            logout.Location = new Point(260, 115);
            logout.Width = 100;
            logout.Click += forgetAuth;
            page.Controls.Add(logout);

            // Status Type
            cmbox = new ComboBox();
            cmbox.DropDownStyle = ComboBoxStyle.DropDownList; // Make it a dropdown. 
            cmbox.Location = new Point(260, 24);
            cmbox.Width = 100;
            cmbox.Items.AddRange(System.Linq.Enumerable.ToArray(statuses.Keys));
            cmbox.SelectedItem = defaultStatus; // I have to take only the names from the dictionary because otherwise I can't select the item I want.
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
        private Boolean setupStatuses()
        {
            if (statuses == null || statuses.ContainsKey("-1"))
            {
                MessageBox.Show("This is a test;");
                statuses = new Dictionary<string, string>();
                JToken stats = request.getStatuses();
                if (checkRequestError() == true)
                {
                    statuses = null;
                    return true;
                }
                    

                foreach (var status in stats)
                {
                    statuses.Add((string)status.SelectToken("text"), (string)status.SelectToken("id"));
                }
            }
            return false;
        }
        
        // Makes sure the communication type is able to be set to Email.
        private Boolean setupCommType()
        {
            if (commType == null || commType == "-1")
            {
                JToken jCommType = request.getCommType();
                if (checkRequestError() == true)
                {
                    commType = null;
                    return true;
                }

                foreach (var cType in jCommType) 
                {
                    if((string)cType.SelectToken("text") == "Email") {
                        commType = (string)cType.SelectToken("id");
                        break;
                    }
                }
            }
            return false;
        }
       
        private void forgetAuth(Object sender, EventArgs e) {
            request.forgetAuth();
            page.Hide();
            MessageBox.Show("You have been logged out.");
        }
       
        private void getAuth()
        {
            login = new Form();
            login.Text = "Login";
            login.Width = 220;
            login.Height = 120;

            userTB = new TextBox();
            userTB.Location = new Point(75,1);
            userTB.Width = 125;
            login.Controls.Add(userTB);

            Label userLabel = new Label();
            userLabel.Text = "Email:";
            userLabel.Width = 75;
            userLabel.Location = new Point(1, 3);
            login.Controls.Add(userLabel);


            passTB = new TextBox();
            passTB.Location = new Point(75, 25);
            passTB.Width = 125;
            passTB.UseSystemPasswordChar = true;
            login.Controls.Add(passTB);

            Label passLabel = new Label();
            passLabel.Text = "Password:";
            passLabel.Width = 75;
            passLabel.Location = new Point(1, 28);
            login.Controls.Add(passLabel);

            Button loginButton = new Button();
            loginButton.Text = "Login";
            loginButton.Location = new Point(150, 50);
            loginButton.Width = 50;
            loginButton.Click += saveAuth;
            login.Controls.Add(loginButton);

            Label messageLabel = new Label();
            messageLabel.Text = "Please login and try your action again.";
            messageLabel.Width = 150;
            messageLabel.Height = 30;
            messageLabel.Location = new Point(1, 50);
            login.Controls.Add(messageLabel);

            login.Show();
        }

        private void saveAuth(Object sender, EventArgs e)
        {
            login.Hide();
            request.getToken(userTB.Text, passTB.Text);
            checkRequestError();            
        }

        private Boolean searchEmail() {
            JToken json = request.searchProjects(mail.SenderEmailAddress);
            if (checkRequestError() == true) return true;

            updateButtons(json);
            return false;
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
            checkRequestError(); // No error checking because it wouldn't go anywhere anyway. 

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
            checkRequestError(); // No error checking because it wouldn't go anywhere anyway. 

            page.Hide();


            
            if ((string)status.SelectToken("project_id") == id)
            {
                MessageBox.Show("Email added successfully.");
                // If there is an error, the error handler is responsible for displaying it. 
            }
        
        }
    }
}
