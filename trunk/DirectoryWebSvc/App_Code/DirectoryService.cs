using System;
using System.Web.Services;
using System.Net;
using System.DirectoryServices;
using System.Net.Security;
using System.Xml;
using System.Security.Cryptography.X509Certificates;
using System.Diagnostics;

[WebService(
   // change to wherever you put this on your IIS
   Namespace = "http://yourWebServer/Path/toThis/Project",
    Name = "Google OneBox Directory Web Service",
   Description = "Web Service providing user and user free/busy information")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]

public class DirectoryService : System.Web.Services.WebService
{
    private const string NameSpace = "http://yourWebServer/Path/toThis/Project";
    private const string UserID = "ldap_userID";
    private const string Password = "userID_password";
    private const string Domain = "your_domain";
    private const string MailServer = "mail.yourdomain.com";
    // interval of free/busy schedule in minutes
    public int interval = 15;
    // date integers to be used for start and end date times (8:00 am - 5:00 pm)
    public int yr = Convert.ToInt32(DateTime.Today.Year);
    public int mo = Convert.ToInt32(DateTime.Today.Month);
    public int dy = Convert.ToInt32(DateTime.Today.Day);
    public string queryString = "";
    public int counter = 0;

    // Clean up any resources being used.
    protected override void Dispose(bool disposing)
    {
    }

    [WebMethod]
    public XmlElement GoogleUsers(string query)
    {
        try
        {
            XmlDocument docUserInfo = new XmlDocument();
            DateTime start = new DateTime(yr, mo, dy, 08, 00, 00);
            DateTime end = new DateTime(yr, mo, dy, 22, 00, 00);
            string strFreeBusyInfo = "";
            queryString = query;
            XmlElement elemTemp;

            DirectoryEntry ldap = new DirectoryEntry("LDAP://yourLDAPServer:389/ou=searchPath,dc=keyenergy,dc=com", UserID, Password);
            DirectorySearcher ldapSearcher = new DirectorySearcher(ldap);

            // define ldap search filter using ambiquous name resolution (anr), not disabled where the
            // object class is user and not computer
            ldapSearcher.Filter = "(&(objectclass=user)(!(objectclass=computer)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(&(anr=" + query + "))))";
            ldapSearcher.SearchRoot = ldap;
            ldapSearcher.SearchScope = SearchScope.Subtree;
            ldapSearcher.Sort.Direction = SortDirection.Ascending;

            // Create new XML document
            XmlDeclaration dec = docUserInfo.CreateXmlDeclaration("1.0", "UTF-8", null);
            docUserInfo.AppendChild(dec);

            // Create new OneBoxResults element
            XmlElement elemOneBox = docUserInfo.CreateElement("OneBoxResults", null);
            docUserInfo.AppendChild(elemOneBox);

            // Create the Title Source element
            XmlElement elemtitle = docUserInfo.CreateElement("title", null);

            // Create the urlText element and append it to the Title element
            elemTemp = docUserInfo.CreateElement("urlText", null);
            elemTemp.InnerText = "My Company Employee Directory Results";
            elemtitle.AppendChild(elemTemp);

            // Create the urlLink element and append it to the Title element
            elemTemp = docUserInfo.CreateElement("urlLink", null);
            elemTemp.InnerText = NameSpace + "/DirectoryService.asmx/GoogleUsers?query=" + query;
            elemtitle.AppendChild(elemTemp);
            elemOneBox.AppendChild(elemtitle);

            // Create the IMAGE_SOURCE element
            XmlElement elemImage = docUserInfo.CreateElement("IMAGE_SOURCE", null);
            elemImage.InnerText = "http://yourWebServer/Path/toThis/Project/images/employee.bmp";
            elemOneBox.AppendChild(elemImage);

            
            // loop through each ldap entry from nameSearch
            foreach (System.DirectoryServices.SearchResult result in ldapSearcher.FindAll())
            {
                // we're only returning the first 3 users free/busy schedule so don't waste time getting
                // the schedule on anyone else otherwise the OneBox modules times out.
                // get ldap info on all users though so we can show as additional names
                if (counter < 3)
                {
                    // get free/busy
                    strFreeBusyInfo = GetFreeBusy(FetchResults.GetProperty(result, "mail", ""), start, end, interval);
                }
                // get all users name, email and phone.
                docUserInfo = BuildUserInfoXml(result, docUserInfo, strFreeBusyInfo, start, elemOneBox);

                strFreeBusyInfo = "";
                counter++;
            }
            return docUserInfo.DocumentElement;
        }
        catch (Exception ex)
        {
            XmlDocument docErrorInfo = BuildErrorXml(ex.Message);

            // Return error information.
            return docErrorInfo.DocumentElement;
        }
    }

    private XmlDocument BuildUserInfoXml(SearchResult searchResult, XmlDocument docUserInfo, string strFreeBusyInfo, DateTime start, XmlElement elemOneBox)
    {
        XmlElement elemTemp;
        DateTime dtInterval;
        DateTime dtIntervaladd;
        string dn = FetchResults.GetProperty(searchResult, "distinguishedName", "");
        string firstname = FetchResults.GetProperty(searchResult, "givenName", "");
        string lastname = FetchResults.GetProperty(searchResult, "sn", "");
        string office = FetchResults.GetProperty(searchResult, "physicalDeliveryOfficeName", "");
        string email = FetchResults.GetProperty(searchResult, "mail", "LOWER");
        string phone = FetchResults.GetProperty(searchResult, "telephoneNumber", "");
        string cell = FetchResults.GetProperty(searchResult, "mobile", "");

        // Create the MODULE_RESULT element
        XmlElement elemResults = docUserInfo.CreateElement("MODULE_RESULT", null);

        // Create the Title (firstname + lastname) element and append it to the MODULE_RESULT element
        elemTemp = docUserInfo.CreateElement("Title", null);
        //elemTemp.InnerText = firstname;
        elemTemp.InnerText = firstname + " " + lastname;
        elemResults.AppendChild(elemTemp);

        // Create the FirstName element and append it to the MODULE_RESULT element
        elemTemp = docUserInfo.CreateElement("Field", null);
        elemTemp.SetAttribute("name", "firstname");
        elemTemp.InnerText = firstname;
        elemResults.AppendChild(elemTemp);

        // Create the LastName element and append it to the MODULE_RESULT element
        elemTemp = docUserInfo.CreateElement("Field", null);
        elemTemp.SetAttribute("name", "lastname");
        elemTemp.InnerText = lastname;
        elemResults.AppendChild(elemTemp);

        // Create the Email element and append it to the MODULE_RESULT element
        elemTemp = docUserInfo.CreateElement("Field", null);
        elemTemp.SetAttribute("name", "email");
        elemTemp.InnerText = email;
        elemResults.AppendChild(elemTemp);

        // Create the Office element and append it to the MODULE_RESULT element
        // this data will be used on mouseover of the data
        if (office.Length != 0)
        {
            elemTemp = docUserInfo.CreateElement("Field", null);
            elemTemp.SetAttribute("name", "office");
            elemTemp.InnerText = office;
            elemResults.AppendChild(elemTemp);
        }

        // Create the OfficePhone element and append it to the MODULE_RESULT element
        if (phone.Length != 0)
        {
            elemTemp = docUserInfo.CreateElement("Field", null);
            elemTemp.SetAttribute("name", "officePhone");
            elemTemp.InnerText = phone;
            elemResults.AppendChild(elemTemp);
        }

        // Create the CellPhone element and append it to the MODULE_RESULT element
        if (cell.Length != 0)
        {
            elemTemp = docUserInfo.CreateElement("Field", null);
            elemTemp.SetAttribute("name", "cellPhone");
            elemTemp.InnerText = cell;
            elemResults.AppendChild(elemTemp);
        }

        // Append the MODULE_RESULT element to the root element
        elemOneBox.AppendChild(elemResults);

        // Create the FreeBusy element
        XmlElement elemFreeBusy = docUserInfo.CreateElement("FreeBusy", null);

        // Set interval counter
        dtInterval = start;
        dtIntervaladd = dtInterval.AddMinutes(interval);

        // Populate the FreeBusy element
        for (int i = 0; i < strFreeBusyInfo.Length; i++)
        {
            string strStatus = "";

            switch (strFreeBusyInfo[i].ToString())
            {
                // Free
                case "0":
                    strStatus = "Free";
                    break;

                // Tentative
                case "1":
                    strStatus = "Tentative";
                    break;

                // Busy
                case "2":
                    strStatus = "Busy";
                    break;

                // Out of office
                case "3":
                    strStatus = "OOF";
                    break;

                // Data not available
                default:
                    strStatus = "Free";
                    break;
            }

            elemTemp = docUserInfo.CreateElement("Field", null);
            elemTemp.SetAttribute("name", dtInterval.ToString("HHmm") + "_" + dtIntervaladd.ToString("HHmm"));
            elemTemp.InnerText = strStatus;
            elemResults.AppendChild(elemTemp);

            dtInterval = dtInterval.AddMinutes(interval);
            dtIntervaladd = dtInterval.AddMinutes(interval);

        }

        // Return the user information document
        return docUserInfo;
    }
    //************************************
    // Portions of the GetFreeBusy method by mstehle
    // at http://blogs.msdn.com/mstehle/articles/628573.aspx
    //************************************
    public string GetFreeBusy(string userSMTP, DateTime start,
    DateTime end, int interval)
    {
        string strResponse = "";
        try
        {
            string freebusyURL = string.Format(
                "https://{0}/public/?cmd=freebusy&start={1}&end={2}&interval={3}&u=SMTP:{4}",
                MailServer, Utils.ConvertToISO8601(start, true),
                Utils.ConvertToISO8601(end, true),
                interval.ToString(), userSMTP);

            // accept owa mail server x509 certificate
            ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(AcceptAllCerts);

            // Create the HttpWebRequest object.
            System.Net.HttpWebRequest request = (System.Net.HttpWebRequest)HttpWebRequest.Create(freebusyURL);

            // Create credentials from user name and password
            System.Net.CredentialCache creds = new CredentialCache();
            System.Uri folderUri = new Uri(freebusyURL);
            System.Net.NetworkCredential netCred = new NetworkCredential(UserID, Password, Domain);
            creds.Add(folderUri, "Basic", netCred);

            request.Credentials = creds.GetCredential(folderUri, "Basic");

            // Specify the method
            request.ContentType = "text/xml";
            request.Method = "GET";
            request.KeepAlive = false;
            request.AllowAutoRedirect = false;

            // NOTE: Response will return with HTML not XML if UserAgent isn't set
            request.UserAgent = "Mozilla/4.0(compatible;MSIE 6.0; " +
                "Windows NT 5.1; SV1; .NET CLR 1.1.4322; .NET CLR 2.0.50727; InfoPath.1)";

            // Send the SEARCH method request and get the response from the server
            System.Net.HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            System.IO.Stream responseStream = response.GetResponseStream();

            // read response stream
            System.IO.StreamReader stream = new System.IO.StreamReader(responseStream);

            XmlTextReader reader = null;
            reader = new XmlTextReader(responseStream);

            int flag = 0;

            // read xml stream to parse out free/busy data
            while (reader.Read())
            {
                if (XmlNodeType.Element.ToString().Equals("Element"))
                {
                    if (reader.Name == "a:fbdata")
                    {
                        flag = 1;
                    }
                    if (reader.Name == "" && flag == 1)
                    {
                        flag = 0;
                        // return free/busy value
                        strResponse = reader.Value;
                        break;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            throw ex;
        }
        return strResponse;
    }

    public class Utils
        // ************************************
        // Created By: mstehle
        // Created On: 6/8/06
        // for use with DAV.
        // ************************************
    {   // convert datetime to ISO8601 datetime before running it through the MS exchange server
        public static string ConvertToISO8601(DateTime myDT, bool noMili)
        {
            if (!noMili)
            {
                return myDT.ToString("yyy-MM-ddT") + myDT.Hour + myDT.ToString(":mm:ss.000Z");
            }
            else
            {
                return myDT.ToString("yyy-MM-ddT") + myDT.Hour + myDT.ToString(":mm:ssZ");
            }
        }

        public static string ConvertToISO8601(DateTime myDT)
        {
            return ConvertToISO8601(myDT, false);
        }
    }

    // accepts all the server certificates
    private static bool AcceptAllCerts(object sender, X509Certificate cert, X509Chain chain, SslPolicyErrors sslPolicyErrors)
    {
        return true;
    }

    public class FetchResults
    {   // return string value from ldap search
        public static string GetProperty(System.DirectoryServices.SearchResult searchResult, string PropertyName, string caseType)
        {
            if (searchResult.Properties.Contains(PropertyName))
            {   // return string results in upper, lower or as-is
                switch (caseType)
                {
                    case "LOWER":
                        return searchResult.Properties[PropertyName][0].ToString().ToLower();
                    case "UPPER":
                        return searchResult.Properties[PropertyName][0].ToString().ToUpper();
                    case "":
                        return searchResult.Properties[PropertyName][0].ToString();
                    default:
                        return string.Empty;
                }
            }
            else
            {
                return string.Empty;
            }
        }
    }

    // return error cause to GSA for debugging instead of [ServerError]
    private XmlDocument BuildErrorXml(string strMessage)
    {
        XmlDocument docErrorInfo = new XmlDocument();

        // Create new Xml Document.
        XmlDeclaration dec = docErrorInfo.CreateXmlDeclaration("1.0", null, null);
        docErrorInfo.AppendChild(dec);

        // Create new OneBoxResults element
        XmlElement elemOneBox = docErrorInfo.CreateElement("OneBoxResults", null);
        docErrorInfo.AppendChild(elemOneBox);

        // Create the ResultCode Source element
        XmlElement elemResult = docErrorInfo.CreateElement("resultCode", null);
        elemResult.InnerText = "WebSvc Error";
        elemOneBox.AppendChild(elemResult);

        // Create the Diagnostics Source element
        XmlElement elemDiagnostic = docErrorInfo.CreateElement("Diagnostics", null);
        elemDiagnostic.InnerText = strMessage;
        elemOneBox.AppendChild(elemDiagnostic);

        return docErrorInfo;
    }
}
