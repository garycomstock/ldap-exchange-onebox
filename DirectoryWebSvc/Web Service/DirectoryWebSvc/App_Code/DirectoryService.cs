using System;
using System.Web.Services;
using System.Net;
using System.DirectoryServices;
using System.Xml;
using System.Configuration;
using System.IO;

[WebService(
   // change to wherever you put this on your IIS
   Namespace = "http://yourWebServer",
    Name = "Google OneBox Directory Web Service",
   Description = "Web Service providing user and user free/busy information")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]

public class DirectoryService : System.Web.Services.WebService
{
    // Web.Config values
    private string authType = ConfigurationManager.AppSettings["authType"];
    private string ldapServer = ConfigurationManager.AppSettings["ldapServer"];
    private string owaURL = ConfigurationManager.AppSettings["owaURL"];
    private string userID = ConfigurationManager.AppSettings["userID"];
    private string domain = ConfigurationManager.AppSettings["domain"];
    private string password = ConfigurationManager.AppSettings["password"];
    private string image = ConfigurationManager.AppSettings["image"];
    private string nameSpace = ConfigurationManager.AppSettings["namespace"];
    // interval of free/busy schedule in minutes
    public int interval = 15;
    // date integers to be used for free/busy start and end date times (8:00 am - 5:00 pm)
    public int yr = Convert.ToInt32(DateTime.Today.Year);
    public int mo = Convert.ToInt32(DateTime.Today.Month);
    public int dy = Convert.ToInt32(DateTime.Today.Day);
    public string queryString = "";
    public int counter = 0;
    public XmlElement elemTemp;
    // clean up any resources being used.
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
            
            // create ldap object using parameters from web.config
            DirectoryEntry ldap = new DirectoryEntry(ldapServer, userID, password);
            DirectorySearcher ldapSearcher = new DirectorySearcher(ldap);

            // define ldap search filter using ambiquous name resolution (anr), user is not disabled and the
            // object class is user and not computer
            ldapSearcher.Filter = "(&(objectclass=user)(!(objectclass=computer)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(&(anr=" + query + "))))";
            ldapSearcher.SearchRoot = ldap;
            ldapSearcher.SearchScope = SearchScope.Subtree;
            ldapSearcher.Sort.Direction = SortDirection.Ascending;
            ldapSearcher.Sort.PropertyName = "givenname";

            // create new XML document
            XmlDeclaration dec = docUserInfo.CreateXmlDeclaration("1.0", "UTF-8", null);
            docUserInfo.AppendChild(dec);

            // create new OneBoxResults element
            XmlElement elemOneBox = docUserInfo.CreateElement("OneBoxResults", null);
            docUserInfo.AppendChild(elemOneBox);

            // create the Title Source element
            XmlElement elemtitle = docUserInfo.CreateElement("title", null);

            // create the urlText element and append it to the Title element
            elemTemp = docUserInfo.CreateElement("urlText", null);
            elemTemp.InnerText = "Employee Directory Results";
            elemtitle.AppendChild(elemTemp);

            // create the urlLink element and append it to the Title element
            elemTemp = docUserInfo.CreateElement("urlLink", null);
            elemTemp.InnerText = nameSpace + "/DirectoryService.asmx/GoogleUsers?query=" + query;
            elemtitle.AppendChild(elemTemp);
            elemOneBox.AppendChild(elemtitle);

            // create the IMAGE_SOURCE element
            XmlElement elemImage = docUserInfo.CreateElement("IMAGE_SOURCE", null);
            elemImage.InnerText = nameSpace + "images/" + image;
            elemOneBox.AppendChild(elemImage);

            // loop through each ldap entry from nameSearch
            foreach (System.DirectoryServices.SearchResult result in ldapSearcher.FindAll())
            {
                // the ldap search for (example) "smith" may return 25 users in your organization and this web
                // service will return them all but we're only going to return the first 3 users free/busy schedule
                // otherwise the OneBox module can time out if a large number of users are returned by the web service.
                // we'll show the ldap info on all users returned by the search so we can show as additional names
                if (counter < 3)
                {
                    // get free/busy information
                    strFreeBusyInfo = GetFreeBusy(FetchResults.GetProperty(result, "mail", ""), start, end, interval);
                }
                // get all users name, email, office and phone(s) that are returned by ldap search
                docUserInfo = BuildUserInfoXml(result, docUserInfo, strFreeBusyInfo, start, elemOneBox);

                // reset free/busy string for next user and increment counter
                strFreeBusyInfo = "";
                counter++;
            }
            return docUserInfo.DocumentElement;
        }
        catch (Exception ex)
        {
            XmlDocument docErrorInfo = BuildErrorXml(ex.Message);

            // return error information.
            return docErrorInfo.DocumentElement;
        }
    }

    private XmlDocument BuildUserInfoXml(SearchResult searchResult, XmlDocument docUserInfo, string strFreeBusyInfo, DateTime start, XmlElement elemOneBox)
    {
        DateTime dtInterval;
        DateTime dtIntervaladd;
        string dn = FetchResults.GetProperty(searchResult, "distinguishedName", "");
        string firstname = FetchResults.GetProperty(searchResult, "givenName", "");
        string lastname = FetchResults.GetProperty(searchResult, "sn", "");
        string office = FetchResults.GetProperty(searchResult, "physicalDeliveryOfficeName", "");
        string email = FetchResults.GetProperty(searchResult, "mail", "LOWER");
        string phone = FetchResults.GetProperty(searchResult, "telephoneNumber", "");
        string cell = FetchResults.GetProperty(searchResult, "mobile", "");

        // create the MODULE_RESULT element
        XmlElement elemResults = docUserInfo.CreateElement("MODULE_RESULT", null);

        // create the Title (firstname + lastname) element and append it to the MODULE_RESULT element
        elemTemp = docUserInfo.CreateElement("Title", null);
        elemTemp.InnerText = firstname + " " + lastname;
        elemResults.AppendChild(elemTemp);

        // create the FirstName element and append it to the MODULE_RESULT element
        elemTemp = docUserInfo.CreateElement("Field", null);
        elemTemp.SetAttribute("name", "firstname");
        elemTemp.InnerText = firstname;
        elemResults.AppendChild(elemTemp);

        // create the LastName element and append it to the MODULE_RESULT element
        elemTemp = docUserInfo.CreateElement("Field", null);
        elemTemp.SetAttribute("name", "lastname");
        elemTemp.InnerText = lastname;
        elemResults.AppendChild(elemTemp);

        // create the Email element and append it to the MODULE_RESULT element
        elemTemp = docUserInfo.CreateElement("Field", null);
        elemTemp.SetAttribute("name", "email");
        elemTemp.InnerText = email;
        elemResults.AppendChild(elemTemp);

        // create the Office element and append it to the MODULE_RESULT element
        // this data will be used on mouseover of the data
        if (office.Length != 0)
        {
            elemTemp = docUserInfo.CreateElement("Field", null);
            elemTemp.SetAttribute("name", "office");
            elemTemp.InnerText = office;
            elemResults.AppendChild(elemTemp);
        }

        // create the OfficePhone element and append it to the MODULE_RESULT element
        if (phone.Length != 0)
        {
            elemTemp = docUserInfo.CreateElement("Field", null);
            elemTemp.SetAttribute("name", "officePhone");
            elemTemp.InnerText = phone;
            elemResults.AppendChild(elemTemp);
        }

        // create the CellPhone element and append it to the MODULE_RESULT element
        if (cell.Length != 0)
        {
            elemTemp = docUserInfo.CreateElement("Field", null);
            elemTemp.SetAttribute("name", "cellPhone");
            elemTemp.InnerText = cell;
            elemResults.AppendChild(elemTemp);
        }

        // Append the MODULE_RESULT element to the root element
        elemOneBox.AppendChild(elemResults);

        // create the FreeBusy element
        XmlElement elemFreeBusy = docUserInfo.CreateElement("FreeBusy", null);

        // Set interval counter
        dtInterval = start;
        dtIntervaladd = dtInterval.AddMinutes(interval);

        // populate the FreeBusy element
        for (int i = 0; i < strFreeBusyInfo.Length; i++)
        {
            string strStatus = "";

            switch (strFreeBusyInfo[i].ToString())
            {
                // free
                case "0":
                    strStatus = "Free";
                    break;

                // tentative
                case "1":
                    strStatus = "Tentative";
                    break;

                // busy
                case "2":
                    strStatus = "Busy";
                    break;

                // out of office
                case "3":
                    strStatus = "OOF";
                    break;

                // data not available
                default:
                    strStatus = "Free";
                    break;
            }

            // create the time interval element i.e.
            // <Field name="0800_0815">Free</Field>
            // <Field name="0815_0830">Busy</Field>
            elemTemp = docUserInfo.CreateElement("Field", null);
            elemTemp.SetAttribute("name", dtInterval.ToString("HHmm") + "_" + dtIntervaladd.ToString("HHmm"));
            elemTemp.InnerText = strStatus;
            elemResults.AppendChild(elemTemp);

            // increment the time intevals by 15 minutes for the next free/busy time
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
    public string GetFreeBusy(string userSMTP, DateTime start, DateTime end, int interval)
    {
        string strResponse = "";

        // if there is an email address for this LDAP user
        if (userSMTP != "")
        {

            Authorization auth = new Authorization();
            HttpWebResponse response;
            string cookies;
            try
            {
                string freebusyURL = string.Format(
                    "{0}/public/?cmd=freebusy&start={1}&end={2}&interval={3}&u=SMTP:{4}",
                    owaURL, Utils.ConvertToISO8601(start, true),
                    Utils.ConvertToISO8601(end, true),
                    interval.ToString(), userSMTP);

                Uri freeBusy = new Uri(freebusyURL);

                // authenticate to owa using method indicated in web.config
                if (authType == "forms")
                {
                    // get cookies from forms auth login session
                    cookies = auth.GetFormsAuthCookies(owaURL, domain, userID, password);
                    // use cookies from prior request to get free/busy url httpwebresponse
                    response = auth.OWAFormsAuthentication(freeBusy, cookies);
                }
                else
                {
                    // httpwebresponse from basic authentication httpwebrequest
                    response = auth.OWABasicAuthentication(freeBusy, userID, password, domain);
                }

                // read response stream from response object
                Stream responseStream = response.GetResponseStream();

                // read response stream
                StreamReader stream = new System.IO.StreamReader(responseStream);

                XmlTextReader reader = null;
                reader = new XmlTextReader(responseStream);

                // read xml stream to parse out free/busy data
                while (reader.Read())
                {
                    if (reader.Name == "a:fbdata")
                    {
                        strResponse = reader.ReadElementString();
                        return strResponse;
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        return strResponse;
    }

    // return error cause to GSA for debugging instead of [ServerError]
    private XmlDocument BuildErrorXml(string strMessage)
    {
        XmlDocument docErrorInfo = new XmlDocument();

        // create new Xml Document.
        XmlDeclaration dec = docErrorInfo.CreateXmlDeclaration("1.0", null, null);
        docErrorInfo.AppendChild(dec);

        // create new OneBoxResults element
        XmlElement elemOneBox = docErrorInfo.CreateElement("OneBoxResults", null);
        docErrorInfo.AppendChild(elemOneBox);

        // create the ResultCode Source element
        XmlElement elemResult = docErrorInfo.CreateElement("resultCode", null);
        elemResult.InnerText = "WebSvc Error";
        elemOneBox.AppendChild(elemResult);

        // create the Diagnostics Source element
        XmlElement elemDiagnostic = docErrorInfo.CreateElement("Diagnostics", null);
        elemDiagnostic.InnerText = strMessage;
        elemOneBox.AppendChild(elemDiagnostic);

        return docErrorInfo;
    }

    public class Utils
    // ************************************
    // created By: mstehle
    // created On: 6/8/06
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
}
