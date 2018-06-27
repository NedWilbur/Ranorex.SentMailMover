// Name:        SentMailMover
// Created By:  Ned Wilbur
// Start Date:  04-01-17
// Desc:        Automatically moves sent items from {mailfilter} to appropriate box
// Usage:       Install and forget it.

using System;
using System.IO;
using System.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SentMailMover
{
    public partial class ThisAddIn
    {
        //Settings
        bool debug = false;

        string mailFilter = "[SenderName] = 'Ranorex Support US' or " +
                            "[SenderName] = 'Ranorex Enterprise Support US' or " +
                            "[SenderName] = 'Ranorex Support' or " +
                            "[SenderName] = 'Enterprise Support' or" +
                            "[SentOnBehalfOfName] = 'Ranorex Support US' or" +
                            "[SentOnBehalfOfName] = 'Ranorex Support'";

        //string mailFilter = "[SenderName] = 'Ned Wilbur'";

        //Variables
        Outlook.NameSpace ns;
        Outlook.MAPIFolder pSentBox, rx_usSentBox, rx_atSentBox;
        Outlook.Items pItems;

        /// <summary>
        /// Outlook Startup
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //Create Required Variables
            try
            {
                log("------------------------------------------------------------------");
                log("Loading SentMailMover");
                
                ns = Application.GetNamespace("MAPI");
                log($"Namespace loaded: {ns.CurrentUser.Name}");

                pSentBox = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail);
                log($"Personal sentbox loaded: {pSentBox.FolderPath}");

                rx_usSentBox = ns.Folders["Ranorex Support US"].Folders["Sent Items"];
                log($"RxUS sentbox loaded: {rx_usSentBox.FolderPath}");

                rx_atSentBox = ns.Folders["Ranorex Support"].Folders["Sent Items"];
                log($"RxAT sentbox loaded: {rx_atSentBox.FolderPath}");

                pItems = pSentBox.Items.Restrict(mailFilter);
                log($"Personal sentbox items: {pItems.Count} (with Filter Applied='{mailFilter}')");
            }
            catch (Exception ex)
            {
                log("Unable to create required variables");
                log($"Exception Thrown: {ex}");
                throw;
            }

            //Create NewItem Handler
            try
            {
                pItems.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(items_ItemAdd);
                log("ItemAdd event listener attached to personal sent folder");
            }
            catch (Exception ex)
            {
                log("Unable to create ItemAdd event listner");
                log($"Exception Thrown: {ex}");
                throw;
            }
        }

        /// <summary>
        /// NewItem event handler
        /// </summary>
        /// <param name="item"></param>
        private void items_ItemAdd(object item)
        {
            log("ItemAdd event listener triggered");
            if (item is Outlook.MailItem)
            {
                Outlook.MailItem mailItem = (Outlook.MailItem) item;

                //Move Mail
                try
                {
                    //if (debug) mailItem.UnRead = true;
                    string sender = mailItem.SenderName;
                    string onBehalfOfName = mailItem.SentOnBehalfOfName;
                    log($"Sender: {sender} (OnBehalfOf: {onBehalfOfName}) | SUBJECT: {mailItem.Subject}");

                    //Move mail based on sender (US or AT)
                    if (sender == "Ranorex Support US" ||
                        onBehalfOfName == "Ranorex Support US")
                    {
                        if (!debug) mailItem.Move(rx_usSentBox);
                        log($"MOVED TO: {rx_usSentBox.FolderPath}");
                    }

                    if (sender == "Ranorex Support" ||
                        onBehalfOfName == "Ranorex Support")
                    {
                        if (!debug) mailItem.Move(rx_atSentBox);
                        log($"MOVED TO: {rx_atSentBox.FolderPath}");
                    }
                }
                catch (Exception ex)
                {
                    log($"Exception Thrown: {ex}");
                }
            }
            else
            {
                log($"Item not moved (Not a MailItem)");
            }
            
        }

        /// <summary>
        /// Write to output window/file (%temp%/SentMailMover.log)
        /// </summary>
        /// <param name="logText"></param>
        public void log(string logText)
        {
            //Debug Output
            if (debug) System.Diagnostics.Debug.WriteLine($"[SentMailMover] {logText}");

            //Output to File
            try
            {
                using (StreamWriter writer = new StreamWriter(Path.GetTempPath() + "SentMailMover.log", true))
                    writer.WriteLine($"[{DateTime.Now}]\t {logText}");
            }
            catch (Exception ex)
            {
                if (debug) System.Diagnostics.Debug.WriteLine($"[SentMailMover] Exception Thrown: {ex}");
            }
        }

        #region VSTO generated code

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += new System.EventHandler(ThisAddIn_Startup);
            Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}