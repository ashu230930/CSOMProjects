using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security;
using System.Configuration;
using System.Net;
using System.IO;

namespace NthDayofMonth
{
    class Program
    {
        static void Main(string[] args)
        {
            // DateTime d1 = DateTime.Parse(DateTime.Now.ToShortDateString());
            // DateTime t = utility.NthOf(0, d1.DayOfWeek);
            ////  d1 = DateTime.Parse(DateTime.Now.ToShortDateString()).AddDays(7);
            // t = utility.NthOf(1, d1.DayOfWeek);
            // t = utility.NthOf(2, d1.DayOfWeek);
            // t = utility.NthOf(3, d1.DayOfWeek);
            //  t = utility.NthOf(4, d1.DayOfWeek);
            //  d1 = DateTime.Parse(DateTime.Now.ToShortDateString()).AddDays(14);
            //  t = utility.NthOf(1, d1.DayOfWeek);
            string logFilePath = ConfigurationManager.AppSettings["logFilePath"];

            string logFileName = "GCC-NthDay-log -" + DateTime.Now.Day.ToString() + "-" + DateTime.Now.Month.ToString() + "-" + DateTime.Now.Year.ToString() + ".log";

            using (StreamWriter writer = new StreamWriter(logFilePath + "/" + logFileName, true))
            {
                writer.WriteLine("Data Time:" + DateTime.Now);
                writer.Close();
            }


            DateTime dateValue = DateTime.Now;

            
                ClientContext ctx = spAuthentication(ConfigurationManager.AppSettings["url"]);
                Web web = ctx.Web;
                List sourcelist = ctx.Web.Lists.GetByTitle(ConfigurationManager.AppSettings["Library"]);
                string we = dateValue.DayOfWeek.ToString();
                string caml = @"<View><Query><Where><Eq><FieldRef Name='WeekDay' /><Value Type='Choice'>" + we + "</Value></Eq></Where></Query></View>";
                CamlQuery query = new CamlQuery();
                query.ViewXml = caml;
                ListItemCollection listItems = sourcelist.GetItems(query);
                ctx.Load(listItems);
                ctx.ExecuteQuery();

                //  Console.WriteLine("Total No of Items is:" + listItems.Count);
                try
                {
                foreach (var listItem in listItems)
                {
                    string rec = listItem["Recurrence"].ToString();
                    string wd = listItem["WeekDay"].ToString();
                    string title = listItem["Title"].ToString();
                    double reportTime = Convert.ToDouble(listItem["ReportTime"]);
                    string assignedTo = listItem["AssignedTo"].ToString();
                    string criticality = listItem["Criticality"].ToString();
                    string program = listItem["Program"].ToString();
                    string subProgram = listItem["SubProgram"].ToString();
                    string status = Convert.ToString(listItem["Status"]);
                    string reportCategory = Convert.ToString(listItem["ReportCategory"]);
                    string distro = Convert.ToString(listItem["Distro"]);


                    if (rec != "Last")
                    {
                        int c = int.Parse(rec);
                        DayOfWeek dd = (DayOfWeek)Enum.Parse(typeof(DayOfWeek), wd);
                        bool dt = utility.NthDayOfMonth(dateValue, dd, c);
                        //  Console.WriteLine(dt);
                        if (dt)
                        {
                            List targetlist = ctx.Web.Lists.GetByTitle(ConfigurationManager.AppSettings["DestinationList"]);
                            CamlQuery qry = new CamlQuery();
                            qry.ViewXml = @"<Query><OrderBy><FieldRef Name='Title'  Ascending='TRUE' /></OrderBy></Query>";
                            ListItemCollection listItemss = targetlist.GetItems(qry);
                            ctx.Load(listItemss);
                            ctx.ExecuteQuery();
                            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                            ListItem oListItem = targetlist.AddItem(itemCreateInfo);
                            oListItem["Title"] = title;
                            oListItem["ReportDate"] = DateTime.Now;
                            oListItem["Program"] = program;
                            var dateTimeOnly = DateTime.Now.Date.AddHours(reportTime);
                            oListItem["ReportTime"] = dateTimeOnly;
                            oListItem["AssignedTo"] = assignedTo;
                            oListItem["Status"] = status;
                            oListItem["Criticality"] = criticality;
                            oListItem["SubProgram"] = subProgram;
                            oListItem["TypeofReport1"] = reportCategory;
                            oListItem.Update();
                            ctx.ExecuteQuery();
                            // Console.Write(oListItem.Id);
                            using (var EmailclientContext = ctx)
                            {
                                var emailprpoperties = new EmailProperties();
                                List<string> assignees = assignedTo.Split(';').ToList();
                                emailprpoperties.To = assignees;
                                emailprpoperties.From = "GCCAdmin@247.ai";
                                emailprpoperties.Body = "<b>Good Day!!</b><br/>" + "Please note – the above said task has been assigned to you and is due by " + dateTimeOnly.ToString("MM/dd/yyyy HH:mm") + " hrs, upon completion respond on the same task mail.<br/>In event of delay click <a href=https://247incc.sharepoint.com/sites/GCC/Lists/GCCTransaction/EditForm.aspx?ID=" + oListItem.Id + "> here</a> to update status/ reason" + "<br/>" + "Refer doc #" + distro + "for email distro";
                                emailprpoperties.Subject = title + " ## " + oListItem.Id;

                                Utility.SendEmail(EmailclientContext, emailprpoperties);
                                EmailclientContext.ExecuteQuery();
                            }


                        }
                    }
                    else
                    {
                        DayOfWeek ddl = (DayOfWeek)Enum.Parse(typeof(DayOfWeek), wd);
                        DateTime de1 = utility.GetLastWeekdayOfTheMonth(DateTime.Now, ddl);
                        //  bool dtd = utility.LastDayOfMonth(DateTime.Now);
                        if (de1.Date.ToShortDateString() == dateValue.Date.ToShortDateString())
                        {
                            List targetlist1 = ctx.Web.Lists.GetByTitle("GCCTransaction");
                            CamlQuery qry1 = new CamlQuery();
                            qry1.ViewXml = @"<Query><OrderBy><FieldRef Name='Title'  Ascending='TRUE' /></OrderBy></Query>";
                            ListItemCollection listItemss1 = targetlist1.GetItems(qry1);
                            ctx.Load(listItemss1);
                            ctx.ExecuteQuery();
                            ListItemCreationInformation itemCreateInfo1 = new ListItemCreationInformation();
                            ListItem oListItem1 = targetlist1.AddItem(itemCreateInfo1);
                            oListItem1["Title"] = title;
                            oListItem1["ReportDate"] = DateTime.Now;
                            oListItem1["Program"] = program;
                            var dateTimeOnly1 = DateTime.Now.Date.AddHours(reportTime);
                            oListItem1["ReportTime"] = dateTimeOnly1;
                            oListItem1["AssignedTo"] = assignedTo;
                            oListItem1["Status"] = status;
                            oListItem1["Criticality"] = criticality;
                            oListItem1["SubProgram"] = subProgram;
                            oListItem1["TypeofReport1"] = reportCategory;
                            oListItem1.Update();
                            ctx.ExecuteQuery();
                            //  Console.WriteLine(oListItem1.Id);
                            using (var EmailclientContext1 = ctx)
                            {
                                var emailprpoperties1 = new EmailProperties();
                                List<string> assignees1 = assignedTo.Split(';').ToList();
                                emailprpoperties1.To = assignees1;
                                emailprpoperties1.From = "GCCAdmin@247.ai";
                                emailprpoperties1.Body = "<b>Good Day!!</b><br/>" + "Please note – the above said task has been assigned to you and is due by " + dateTimeOnly1.ToString("MM/dd/yyyy HH:mm") + " hrs, upon completion respond on the same task mail.<br/>In event of delay click <a href=https://247incc.sharepoint.com/sites/GCC/Lists/GCCTransaction/EditForm.aspx?ID=" + oListItem1.Id + "> here</a> to update status/ reason" + "<br/>" + "Refer doc #" + distro + "&nbsp;for email distro";
                                emailprpoperties1.Subject = title + " ## " + oListItem1.Id;

                                Utility.SendEmail(EmailclientContext1, emailprpoperties1);
                                EmailclientContext1.ExecuteQuery();
                            }
                        }
                    }
                }

                using (StreamWriter writer = new StreamWriter(logFilePath + "/" + logFileName, true))
                {
                    writer.WriteLine("Data Time:" + DateTime.Now);
                    writer.Close();
                }

            }
            catch (Exception ex)
            {
                using (var EmailclientContext = ctx)
                {
                    var emailprpoperties = new EmailProperties();
                    emailprpoperties.To = new List<string> { "ashtosh.singh@247.ai", "vinay.chavadi@247.ai" };
                    emailprpoperties.From = "spserviceadmin@247.ai";
                    emailprpoperties.Body = "<b>Nth Day data creation failed</b>" + ex;
                    emailprpoperties.Subject = "Nth Day data creation failed";

                    Utility.SendEmail(EmailclientContext, emailprpoperties);
                    EmailclientContext.ExecuteQuery();
                }
                using (StreamWriter writer = new StreamWriter(logFilePath + "/" + logFileName, true))
                {
                    writer.WriteLine("Exception Msg : "+ ex.Message);
                    writer.WriteLine("Exception Msg : " + ex.StackTrace);
                    writer.Close();
                }
            }
        }


        private static ClientContext spAuthentication(string path)
        {



            ClientContext ctx = new ClientContext(path);
            SecureString secStr = new SecureString();

           
                foreach (var c in ConfigurationManager.AppSettings["Password"].ToCharArray())
                {
                    secStr.AppendChar(c);
                }
            
            ctx.Credentials = new SharePointOnlineCredentials(ConfigurationManager.AppSettings["UserID"], secStr);
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            return ctx;

        }

       




    }
}
