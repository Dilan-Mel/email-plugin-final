using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.IO;
using Newtonsoft.Json;

namespace EmailReadOutlook
{
    public partial class ThisAddIn
    {
        Outlook.NameSpace outlookNameSpace;
        Outlook.MAPIFolder inbox;
        Outlook.Items items;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(ItemSend);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
        }
        private void ItemSend(object Item, ref bool Cancel)
        {
            Outlook.MailItem mail = (Outlook.MailItem)Item;
            if (Item != null)
            {
                if (mail.MessageClass == "IPM.Note")
                {
                    var jsonData = GetText();
                    var inCount = 0;
                    if (mail.Body.Length > 3)
                    {
                        // dynamic json = JsonConvert.DeserializeObject(mail.Body);
                        foreach (var item in jsonData)
                        {
                            if (mail.Body.ToUpper().Contains(item))
                            {
                                inCount += 1;
                            }
                        }

                        if (inCount != 0)
                        {
                            if (inCount > 5)
                            {
                                MessageBox.Show("This email is blocked.please remove any high sensitive information available and resend it.");
                                Cancel = true;
                            }
                            else
                            {
                                var Confirm = MessageBox.Show("Your email contains sensitive information likely to violate privacy policies. Please check again and resend", "Privacy issue", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                                if (Confirm == DialogResult.OK)
                                {
                                    Cancel = false;
                                }
                                else
                                {
                                    Cancel = true;
                                }
                            }
                        }
                    }
                    else
                    {
                        Cancel = true;
                    }
                }
                else Cancel = true;
            }

        }

        private string[] GetText()
        {
            try
            {
                string path = @"C:\\Users\\Dilan Mel\\Documents\\Visual Studio 2015\\Projects\\EmailReadOutlook\\Email-plugin\\EmailReadOutlook\\TextBook.json";

                int inOrOut = 0;
                JsonTextReader reader = new JsonTextReader(new StreamReader(path));
                string stValue = null;
                while (reader.Read())
                {
                    if (reader.Value != null)
                        if (inOrOut == 1)
                        {
                            stValue = reader.Value.ToString().ToUpper();
                            inOrOut = 0;
                        }
                        else
                            inOrOut = 1;
                }

                return stValue.Split(',');
            }
            catch (Exception ex)
            {
                throw ex;
            }
           
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
