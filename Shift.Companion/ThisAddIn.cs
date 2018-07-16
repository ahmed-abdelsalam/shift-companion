using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Shift.Companion
{
    public partial class ThisAddIn

    {
        public static Form1  mainform = new Form1();
        Outlook.NameSpace outlookNameSpace;
        Outlook.MAPIFolder inbox;
        Outlook.Items items;

        public int SearchGroup(int group , string search)
        {
            int index = -1;
            foreach (ListViewItem item in  mainform.customcontrol11.listView1.Groups[group].Items)
            {
                if (item.Text == search)
                    index = item.Index;
            }
            return index;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

            outlookNameSpace = this.Application.GetNamespace("MAPI");
            inbox = outlookNameSpace.GetDefaultFolder(
                    Microsoft.Office.Interop.Outlook.
                    OlDefaultFolders.olFolderInbox);
            items = inbox.Items;
            items.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(InboxFolderItemAdded);


        }


        void InboxFolderItemAdded(object Item)
            {
                
                if (Item is Outlook.MailItem)
                    {

                        Outlook.MailItem mail = (Outlook.MailItem)Item;
                        if (Item != null)
                        {
                            int index;
                            mainform.textBox1.Text = mail.Subject;
                            Regex time = new Regex(@"\b\d{1,2}\:\d{1,2}\ [AaPpMm]{2}");
                            Regex IP = new Regex(@"\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b");
                            



                            if (mail.SenderEmailAddress == "DC-WhatsUp@support.linkdatacenter.net")
                            {
                                Regex word = new Regex(@"\w+");
                                Match mIP = IP.Match(mail.Body);
                                Match service = Regex.Match(mail.Subject, @"(?<before>[A-Za-z0-9,]+|[A-Za-z0-9,]+\s[0-9]+) Down (?<after>\w+)", RegexOptions.IgnoreCase);
                                index = SearchGroup(0, word.Matches(mail.Subject)[1].Value);
                                if (mail.Subject.ToUpper().Contains("Down".ToUpper()))
                                {
                                //Match result2 = time.Match(mail.Subject);
                                    if(index != -1)
                                    {
                                        mainform.customcontrol11.listView1.Items[index].SubItems[2].Text = service.Groups["before"].ToString();
                                    }
                                    else {
                                        ListViewItem li = new ListViewItem(word.Matches(mail.Subject)[1].Value, 1);
                                        li.SubItems.Add(mIP.Value);
                                        li.SubItems.Add(service.Groups["before"].ToString());
                                        li.SubItems.Add("Down");
                                        li.SubItems.Add(mail.ReceivedTime.ToShortTimeString());
                                        li.Group = mainform.customcontrol11.listView1.Groups[0];
                                        mainform.customcontrol11.listView1.Items.Add(li);
                                    }
                                
                                }
                                else if (mail.Subject.ToUpper().Contains("UP".ToUpper()))
                                {
                                    if (index != -1)
                                    {
                                        if(mainform.customcontrol11.listView1.Items[index].SubItems[3].Text== "Down")    
                                            mainform.customcontrol11.listView1.Items[index].Group= mainform.customcontrol11.listView1.Groups[6];
                                    }
                                }
                            }



                            else if (mail.SenderEmailAddress == "sitescope@support.linkdatacenter.net")
                            {
                                
                                var url = Regex.Match(mail.Body, @"(?<=URL: )(.+?)(?=\n|,)");
                                if (mail.Subject.Contains("Sitescope Alert, error"))
                                {

                                
                                    if ( mainform.customcontrol11.listView1.FindItemWithText(url.Value)==null)
                                    {
                                        ListViewItem li = new ListViewItem("Url Down", 3);
                                        li.SubItems.Add(url.Value);
                                        li.SubItems.Add("");
                                        li.SubItems.Add(mail.ReceivedTime.ToShortTimeString());
                                        li.Group = mainform.customcontrol11.listView1.Groups[3];
                                        mainform.customcontrol11.listView1.Items.Add(li);
                                    }
                                }

                                else if (mail.Subject.Contains("Sitescope Alert, good"))
                                {
                                    if (mainform.customcontrol11.listView1.FindItemWithText(url.Value)!= null)
                                    {
                                        mainform.customcontrol11.listView1.FindItemWithText(url.Value).Remove();

                                    }
                                }

                            }





                            else if (mail.SenderEmailAddress == "DC-Solarwinds@support.linkdatacenter.net")
                            {

                                var space = Regex.Match(mail.Subject, @"(?<=Alert: Percent Space Used of )(.+?)(?=-)");
                                var spaceReset = Regex.Match(mail.Subject, @"(?<=Reset: Percent Space Used of )(.+?)(?=-)");
                                var percent = Regex.Match(mail.Subject, @"(?<=is now )(.+?)(?=%)");
                                index = SearchGroup(1, space.Value);
                                if (mail.Subject.Contains("Alert: Percent Space"))
                                {
                                    if (index == -1) {
                                        ListViewItem li = new ListViewItem(space.Value, 0);
                                        li.SubItems.Add(IP.Match(mail.Body).Value);
                                        li.SubItems.Add(percent.Value+"%");
                                        li.SubItems.Add("space");
                                        li.SubItems.Add(mail.ReceivedTime.ToShortTimeString());
                                        li.Group = mainform.customcontrol11.listView1.Groups[1];
                                        mainform.customcontrol11.listView1.Items.Add(li);
                                    }
                                    else
                                    {
                                    mainform.customcontrol11.listView1.Items[index].SubItems[2].Text= percent.Value + "%";
                                    }

                                }


                                else if (mail.Subject.Contains("Reset: Percent Space"))
                                {

                                    if (index != -1)
                                    {
                                        mainform.customcontrol11.listView1.Items[index].Group= mainform.customcontrol11.listView1.Groups[6];
                                        mainform.customcontrol11.listView1.Items[index].SubItems[2].Text= percent.Value + "%";
                                    }               

                                }


                                else if (mail.Subject.Contains("Reboot"))
                                {
                                    var servreName = Regex.Match(mail.Body, @"(?<=ALERT: )(.+?)(?= -)");
                                    




                                }
                            }








                            else if (mail.SenderEmailAddress == "be@support.linkdatacenter.net")
                            {
                                
                                var servreName = Regex.Match(mail.Subject, @"(?<=Server: )(.+?)(?=\))");
                                var jobName = Regex.Match(mail.Subject, @"(?<=Job: )(.+?)(?=\))");
                                    
                                        ListViewItem li = new ListViewItem(servreName.Value,2 );
                                        if (mail.Subject.Contains("Job Failed"))
                                        {
                                            li.SubItems.Add("");
                                            li.SubItems.Add("failed Job");
                                            li.SubItems.Add("Backup");
                                            li.SubItems.Add(jobName.Value);
                                            li.SubItems.Add(mail.ReceivedTime.ToShortTimeString());
                                            li.Group = mainform.customcontrol11.listView1.Groups[7];
                                            mainform.customcontrol11.listView1.Items.Add(li);
                                        }
                                        else if (mail.Subject.Contains("Job Cancellation"))
                                        {
                                            li.SubItems.Add("");
                                            li.SubItems.Add("cancelled Job");
                                            li.SubItems.Add("Backup");
                                            li.SubItems.Add(jobName.Value);
                                            li.SubItems.Add(mail.ReceivedTime.ToShortTimeString());
                                            li.Group = mainform.customcontrol11.listView1.Groups[7];
                                            mainform.customcontrol11.listView1.Items.Add(li);
                                        }
                                        
                   

                            }


                            else if (mail.SenderEmailAddress == "be@support.linkdatacenter.net")
                            {
                                var downName = Regex.Match(mail.Subject, @"(?<=Alarm: )(.+?)(?=is  Down|is  Critical))");
                                var upName = Regex.Match(mail.Subject, @"(?<=Alarm: )(.+?)(?=is  Up|is  Clear))");
                                index = SearchGroup(4, downName.Value);
                                if (mail.Subject.Contains("Down")|| mail.Subject.Contains("Critical"))
                                {if(index==-1)
                                    index = SearchGroup(4, downName.Value);
                                    ListViewItem li = new ListViewItem(downName.Value, 3);
                                    li.SubItems.Add("");
                                    li.SubItems.Add(time.Match(mail.Body).Value);
                                    li.SubItems.Add("Down");
                                    li.Group = mainform.customcontrol11.listView1.Groups[4];
                                    mainform.customcontrol11.listView1.Items.Add(li);

                                }
                                else if( mail.Subject.Contains("Up") || mail.Subject.Contains("Clear"))
                                {
                                    index = SearchGroup(4, upName.Value);
                                    if (index != -1)
                                    {
                                        mainform.customcontrol11.listView1.Items[index].Remove();
                                    }

                                }

                            }






                        }
                    }
            }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
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
