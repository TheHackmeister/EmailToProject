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
        private ProjectRequest request;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
// Probably get Auth token here. 
            request = new ProjectRequest();
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

            new ProjectForm(item, request);           
            //JObject projects = JObject.Parse(json);
            //projects.co
            //Button b = new Button();

      
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
