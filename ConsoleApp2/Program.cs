using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp2
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string country = ConfigurationManager.AppSettings["Country"];
            List<DesearilizacionCountry> list = JsonConvert.DeserializeObject<List<DesearilizacionCountry>>(country);

            foreach (DesearilizacionCountry item in list.ToList())
            {
                getAllUsers(item.OU, item.process +"Users");
                getAllComputers(item.OU, item.process +"Computers");
            }
        }

        static DirectoryEntry createDirectoryEntry(String path)
        {
            DirectoryEntry ldapConnection = new DirectoryEntry(ConfigurationManager.AppSettings["AD_SERVER"], ConfigurationManager.AppSettings["AD_USER"], ConfigurationManager.AppSettings["AD_PASS"]);
            ldapConnection.Path = path;
            ldapConnection.AuthenticationType = AuthenticationTypes.Secure;
            return ldapConnection;
        }
 
        static void getAllUsers(String ou, String process)
        {
            String pathMX = $"LDAP://OU={ou},DC=mlock,DC=com";
            DirectoryEntry myLdapConnection = createDirectoryEntry(pathMX);
            DirectorySearcher search = new DirectorySearcher(myLdapConnection);
            search.Filter = "(&(objectClass=user)(objectCategory=person))";
            search.PropertiesToLoad.Add("cn");
            search.PropertiesToLoad.Add("distinguishedName");
            search.PropertiesToLoad.Add("mail");
            search.PropertiesToLoad.Add("physicalDeliveryOfficeName");
            search.PropertiesToLoad.Add("samaccountname");
            search.PropertiesToLoad.Add("mail");
            search.PropertiesToLoad.Add("userAccountControl");
            search.PropertiesToLoad.Add("canonicalName");
            search.PropertiesToLoad.Add("usergroup");
            search.PropertiesToLoad.Add("displayname");//first name
            search.PropertiesToLoad.Add("lastLogonTimestamp");//last conection
            search.PropertiesToLoad.Add("employeeNumber");

            String cn = "";
            SearchResultCollection resultCol = search.FindAll();

            List<string> listUsers = new List<string>();

            //Console.WriteLine("Usuarios con mas de 30 dias");
            foreach (SearchResult result in resultCol)
            {
                try
                {
                    Int64 time;
                    DateTime? dt = null;
                    DateTime now = DateTime.Now;
                    double totalMilliseconds_TimeStamp, totalMilliseconds_Now, diff;
                    ulong ms_TimeStamp, ms_Now;



                    if (result.Properties.Contains("samaccountname") && result.Properties["userAccountControl"][0].ToString() == "512" && result.Properties["canonicalName"][0].ToString().Contains("Users"))
                    //result.Properties["samaccountname"][0].ToString() == "vcamez" o result.Properties.Contains("samaccountname") && result.Properties.Contains("employeeNumber")
                    //result.Properties.Contains("samaccountname") && result.Properties.Contains("employeeNumber") && result.Properties["userAccountControl"][0].ToString() == "512"
                    {
                        if (result.Properties.Contains("lastLogonTimestamp"))
                        {
                            time = Convert.ToInt64(result.Properties["lastLogonTimestamp"][0].ToString());
                            dt = new DateTime(1601, 01, 01, 0, 0, 0, DateTimeKind.Utc).AddTicks(time);

                            DateTime dt2 = dt.Value;

                            totalMilliseconds_TimeStamp = dt2.ToUniversalTime().Subtract(new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds;
                            ms_TimeStamp = (ulong)totalMilliseconds_TimeStamp;
                            totalMilliseconds_Now = now.ToUniversalTime().Subtract(new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds;
                            ms_Now = (ulong)totalMilliseconds_Now;

                            TimeSpan ts = now.Subtract(dt2);
                            diff = ts.TotalDays;

                            if (diff >= 30)
                            {

                                //Console.WriteLine($"Name: {result.Properties["cn"][0].ToString()}; Last logon: {(dt.HasValue ? dt.ToString() : "")}");
                                //DisableCNObjects(result.Properties["cn"][0].ToString());
                                listUsers.Add("<b>" + result.Properties["cn"][0].ToString() + "</b>" + " - <i>" + dt.ToString() + "</i>");

                            }
                        }
                        else
                        {
                            //Console.WriteLine($"Name: {result.Properties["cn"][0].ToString()}; Last logon: {(dt.HasValue ? dt.ToString() : "")}");
                            //DisableCNObjects(result.Properties["cn"][0].ToString());
                            listUsers.Add("<b>" + result.Properties["cn"][0].ToString() + "</b>" + " - <i> No date </i>");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }

            if (listUsers.Count == 0)
            {
                //Console.WriteLine("No se encontro a ningun usuario en espana");
            }
            else
            {
                IT_NotificationAD(process, listUsers);
            }

        }

        static void getAllComputers(String ou, String process)
        {
            String pathMX = $"LDAP://OU={ou},DC=mlock,DC=com";
            DirectoryEntry myLdapConnection = createDirectoryEntry(pathMX);
            DirectorySearcher search = new DirectorySearcher(myLdapConnection);
            search.Filter = "((objectCategory=computer))";
            search.PropertiesToLoad.Add("cn");
            search.PropertiesToLoad.Add("samaccountname");
            search.PropertiesToLoad.Add("canonicalName");
            search.PropertiesToLoad.Add("lastLogonTimestamp");//last conection

            String cn = "";
            SearchResultCollection resultCol = search.FindAll();

            List<String> listPc = new List<string>();

            foreach (SearchResult result in resultCol)
            {
                try
                {
                    Int64 time;
                    DateTime? dt = null;
                    DateTime now = DateTime.Now;
                    double totalMilliseconds_TimeStamp, totalMilliseconds_Now, diff;
                    ulong ms_TimeStamp, ms_Now;
                    List<string> computers = new List<string>();

                    if (result.Properties.Contains("samaccountname") && result.Properties["canonicalName"][0].ToString().Contains("Workstations"))
                    //result.Properties["samaccountname"][0].ToString() == "MXMLPF3HMBNS" o result.Properties.Contains("samaccountname")
                    {
                        if (result.Properties.Contains("lastLogonTimestamp"))
                        {

                            cn = result.Properties["cn"][0].ToString();
                            time = Convert.ToInt64(result.Properties["lastLogonTimestamp"][0].ToString());
                            dt = new DateTime(1601, 01, 01, 0, 0, 0, DateTimeKind.Utc).AddTicks(time);

                            DateTime dt2 = dt.Value;

                            totalMilliseconds_TimeStamp = dt2.ToUniversalTime().Subtract(new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds;
                            ms_TimeStamp = (ulong)totalMilliseconds_TimeStamp;
                            totalMilliseconds_Now = now.ToUniversalTime().Subtract(new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds;
                            ms_Now = (ulong)totalMilliseconds_Now;

                            TimeSpan ts = now.Subtract(dt2);
                            diff = ts.TotalDays;

                            if (diff >= 30)
                            {
                                //Console.WriteLine($"Name: {result.Properties["cn"][0].ToString()}; Last logon: {(dt.HasValue ? dt.ToString() : "")}");  
                                listPc.Add("<b>" + result.Properties["cn"][0].ToString() + "</b>" + " - <i>" + dt.ToString() + "</i>");
                            }

                        }
                        else
                        {
                            //Console.WriteLine($"Name: {result.Properties["cn"][0].ToString()}; Last logon: {(dt.HasValue ? dt.ToString() : "")}");
                            listPc.Add("<b>" + result.Properties["cn"][0].ToString() + "</b>" + " - <i> No date </i>");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }

            if (listPc.Count == 0)
            {
                //Console.WriteLine("No se encontro a ningun usuario en espana");
            }
            else
            {
                IT_NotificationAD(process, listPc);
            }


        }

        private static void DisableCNObjects(String cn)
        {
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain, ConfigurationManager.AppSettings["AD_SERVER"], ConfigurationManager.AppSettings["AD_USER"], ConfigurationManager.AppSettings["AD_PASS"]);
            UserPrincipal userPrincipal = UserPrincipal.FindByIdentity(ctx, cn);
            userPrincipal.Enabled = false;
            userPrincipal.Save();

            ctx.Dispose();
        }
        static void IT_NotificationAD(String process, List<string> list)
        {
            string url = @"https://fbos.service-now.com/sp";
            MailMessage mail = new MailMessage();

            mail.IsBodyHtml = true;
            mail.From = new MailAddress("Tress-AD@mlock.com");
            mail.To.Add(ConfigurationManager.AppSettings["IT_Notification"]); //fbosdev@service-now.com jcamacho2@mlock.com
            StringBuilder Body = new StringBuilder();
            switch (process)
            {
                case "ReviewMXUsers":
                    mail.Subject = "(MEXICO) Users without activity (30 days)";
                    Body.Append("Please check the following users, they have not logged in for more than 30 days or have no usage: ");
                    Body.Append($"<ul>");
                    foreach (string item in list)
                    {
                        Body.Append($"<li>{item}</li>");
                    }
                    Body.Append($"</ul>");
                    Body.Append("Have been disabled");
                    break;

                case "ReviewMXComputers":
                    mail.Subject = "(MEXICO) Computers without activity (30 days)";

                    Body.Append("Please Review the next computers without activity for more than 30 days: <br/>");
                    Body.Append($"<ul>");
                    foreach (string item in list)
                    {
                        Body.Append($"<li>{item}</li>");
                    }
                    Body.Append($"</ul>");
                    break;
                case "ReviewUSAUsers":
                    mail.Subject = "(USA) Users without activity (30 days)";
                    Body.Append("Please check the following users, they have not logged in for more than 30 days or have no usage: ");
                    Body.Append($"<ul>");
                    foreach (string item in list)
                    {
                        Body.Append($"<li>{item} </li>");
                    }
                    Body.Append($"</ul>");
                    Body.Append("Have been disabled");
                    break;
                case "ReviewUSAComputers":
                    mail.Subject = "(USA) Computers without activity (30 days)";

                    Body.Append("Please Review the next computers without activity for more than 30 days: <br/>");
                    Body.Append($"<ul>");
                    foreach (string item in list)
                    {
                        Body.Append($"<li>{item}</li>");
                    }
                    Body.Append($"</ul>");
                    break;
                case "ReviewCAUsers":
                    mail.Subject = "(CANADA) Users without activity (30 days)";
                    Body.Append("Please check the following users, they have not logged in for more than 30 days or have no usage: ");
                    Body.Append($"<ul>");
                    foreach (string item in list)
                    {
                        Body.Append($"<li>{item} </li>");
                    }
                    Body.Append($"</ul>");
                    Body.Append("Have been disabled");
                    break;
                case "ReviewCAComputers":
                    mail.Subject = "(CANADA) Computers without activity (30 days)";

                    Body.Append("Please Review the next computers without activity for more than 30 days: <br/>");
                    Body.Append($"<ul>");
                    foreach (string item in list)
                    {
                        Body.Append($"<li>{item}</li>");
                    }
                    Body.Append($"</ul>");
                    break;
                case "ReviewCNUsers":
                    mail.Subject = "(CHINA) Users without activity (30 days)";
                    Body.Append("Please check the following users, they have not logged in for more than 30 days or have no usage: ");
                    Body.Append($"<ul>");
                    foreach (string item in list)
                    {
                        Body.Append($"<li>{item} </li>");
                    }
                    Body.Append($"</ul>");
                    Body.Append("Have been disabled");
                    break;
                case "ReviewCNComputers":
                    mail.Subject = "(CHINA) Computers without activity (30 days)";

                    Body.Append("Please Review the next computers without activity for more than 30 days: <br/>");
                    Body.Append($"<ul>");
                    foreach (string item in list)
                    {
                        Body.Append($"<li>{item}</li>");
                    }
                    Body.Append($"</ul>");
                    break;
                case "ReviewDEUsers":
                    mail.Subject = "(GERMANY) Users without activity (30 days)";
                    Body.Append("Please check the following users, they have not logged in for more than 30 days or have no usage: ");
                    Body.Append($"<ul>");
                    foreach (string item in list)
                    {
                        Body.Append($"<li>{item} </li>");
                    }
                    Body.Append($"</ul>");
                    Body.Append("Have been disabled");
                    break;
                case "ReviewDEComputers":
                    mail.Subject = "(GERMANY) Computers without activity (30 days)";

                    Body.Append("Please Review the next computers without activity for more than 30 days: <br/>");
                    Body.Append($"<ul>");
                    foreach (string item in list)
                    {
                        Body.Append($"<li>{item}</li>");
                    }
                    Body.Append($"</ul>");
                    break;
                case "ReviewENUsers":
                    mail.Subject = "(ENGLAND) Users without activity (30 days)";
                    Body.Append("Please check the following users, they have not logged in for more than 30 days or have no usage: ");
                    Body.Append($"<ul>");
                    foreach (string item in list)
                    {
                        Body.Append($"<li>{item} </li>");
                    }
                    Body.Append($"</ul>");
                    Body.Append("Have been disabled");
                    break;
                case "ReviewENComputers":
                    mail.Subject = "(ENGLAND) Computers without activity (30 days)";

                    Body.Append("Please Review the next computers without activity for more than 30 days: <br/>");
                    Body.Append($"<ul>");
                    foreach (string item in list)
                    {
                        Body.Append($"<li>{item}</li>");
                    }
                    Body.Append($"</ul>");
                    break;
                case "ReviewFRUsers":
                    mail.Subject = "(FRANCE) Users without activity (30 days)";
                    Body.Append("Please check the following users, they have not logged in for more than 30 days or have no usage: ");
                    Body.Append($"<ul>");
                    foreach (string item in list)
                    {
                        Body.Append($"<li>{item} </li>");
                    }
                    Body.Append($"</ul>");
                    Body.Append("Have been disabled");
                    break;
                case "ReviewFRComputers":
                    mail.Subject = "(FRANCE) Computers without activity (30 days)";

                    Body.Append("Please Review the next computers without activity for more than 30 days: <br/>");
                    Body.Append($"<ul>");
                    foreach (string item in list)
                    {
                        Body.Append($"<li>{item}</li>");
                    }
                    Body.Append($"</ul>");
                    break;
                case "ReviewJPUsers":
                    mail.Subject = "(JAPAN) Users without activity (30 days)";
                    Body.Append("Please check the following users, they have not logged in for more than 30 days or have no usage: ");
                    Body.Append($"<ul>");
                    foreach (string item in list)
                    {
                        Body.Append($"<li>{item} </li>");
                    }
                    Body.Append($"</ul>");
                    Body.Append("Have been disabled");
                    break;
                case "ReviewJPComputers":
                    mail.Subject = "(JAPAN) Computers without activity (30 days)";

                    Body.Append("Please Review the next computers without activity for more than 30 days: <br/>");
                    Body.Append($"<ul>");
                    foreach (string item in list)
                    {
                        Body.Append($"<li>{item}</li>");
                    }
                    Body.Append($"</ul>");
                    break;
                case "ReviewNLUsers":
                    mail.Subject = "(NETHERLANDS) Users without activity (30 days)";
                    Body.Append("Please check the following users, they have not logged in for more than 30 days or have no usage: ");
                    Body.Append($"<ul>");
                    foreach (string item in list)
                    {
                        Body.Append($"<li>{item} </li>");
                    }
                    Body.Append($"</ul>");
                    Body.Append("Have been disabled");
                    break;
                case "ReviewNLComputers":
                    mail.Subject = "(NETHERLANDS) Computers without activity (30 days)";

                    Body.Append("Please Review the next computers without activity for more than 30 days: <br/>");
                    Body.Append($"<ul>");
                    foreach (string item in list)
                    {
                        Body.Append($"<li>{item}</li>");
                    }
                    Body.Append($"</ul>");
                    break;
                case "ReviewSPUsers":
                    mail.Subject = "(SPAIN) Users without activity (30 days)";
                    Body.Append("Please check the following users, they have not logged in for more than 30 days or have no usage: ");
                    Body.Append($"<ul>");
                    foreach (string item in list)
                    {
                        Body.Append($"<li>{item} </li>");
                    }
                    Body.Append($"</ul>");
                    Body.Append("Have been disabled");
                    break;
                case "ReviewSPComputers":
                    mail.Subject = "(SPAIN) Computers without activity (30 days)";

                    Body.Append("Please Review the next computers without activity for more than 30 days: <br/>");
                    Body.Append($"<ul>");
                    foreach (string item in list)
                    {
                        Body.Append($"<li>{item}</li>");
                    }
                    Body.Append($"</ul>");
                    break;

                default:
                    Console.WriteLine("Mensaje no enviado");
                    break;



            }
            Body.Append("<br/> ");
            Body.Append("<font size=1>this message is created automated, if you have any issues please contact IT Help desk 3709 or go to <b> <a href=" + url + ">ServiceNow</a></b> to submit a support ticket </font><br/> ");

            mail.Body = Body.ToString();
            mail.IsBodyHtml = true;
            SmtpClient smtp = new SmtpClient("NMMailRelay01.mlock.com", 25);
            smtp.Send(mail);
        }

    }
    public class DesearilizacionCountry
    {
        public string OU { get; set; }
        public string process { get; set; }

    }


}
