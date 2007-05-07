using System;
using System.Configuration;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using System.Text.RegularExpressions;
using System.Net;
using System.Text;
using System.IO;

/// <summary>
/// Summary description for Auth
/// </summary>
public class Authorization
{
    public string owaURL = ConfigurationManager.AppSettings["owaURL"];
    public string proxyURL = ConfigurationManager.AppSettings["proxyURL"];

    public HttpWebResponse OWABasicAuthentication(Uri freebusyURL, string userID, string password, string domain)
    {
        // accept owa mail server x509 certificate using a callback
        ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(AcceptAllCerts);

        // create free/busy http web request to the owa freebusy url
        HttpWebRequest fbRequest = (HttpWebRequest)WebRequest.Create(freebusyURL);

        // create credentials from user name, password and domain
        CredentialCache creds = new CredentialCache();
        NetworkCredential netCred = new NetworkCredential(userID, password, domain);
        creds.Add(freebusyURL, "Basic", netCred);
        
        // add credentials to http web request
        fbRequest.Credentials = creds.GetCredential(freebusyURL, "Basic");

        // Specify add'l header info
        fbRequest.ContentType = "text/xml";
        fbRequest.Method = "GET";
        fbRequest.KeepAlive = false;
        fbRequest.AllowAutoRedirect = false;

        // NOTE: Response will return with HTML not XML if UserAgent isn't set
        fbRequest.UserAgent = "Mozilla/4.0(compatible;MSIE 6.0; " +
            "Windows NT 5.1; SV1; .NET CLR 1.1.4322; .NET CLR 2.0.50727; InfoPath.1)";

        // if proxy is used then add proxy url to http request
        if (proxyURL != "")
        {
            WebProxy webProxy = new WebProxy();
            Uri proxyAddr = new Uri(proxyURL);
            webProxy.Address = proxyAddr;
            fbRequest.Proxy = webProxy;
        }

        // response will return with HTML not XML if UserAgent isn't set
        fbRequest.UserAgent = "Mozilla/4.0(compatible;MSIE 6.0; " +
            "Windows NT 5.1; SV1; .NET CLR 1.1.4322; .NET CLR 2.0.50727; InfoPath.1)";

        // return the response from http web request
        return (HttpWebResponse)fbRequest.GetResponse();
    }

    public HttpWebResponse OWAFormsAuthentication(Uri freebusyURL, string formsAuthCookie)
    {
        // create free/busy http web request to the owa freebusy url
        HttpWebRequest fbRequest = (HttpWebRequest)WebRequest.Create(freebusyURL);
        
        // add cookies from prior authenticated session to new request header
        fbRequest.Headers.Add("Cookie", formsAuthCookie);

        // if proxy is used then add proxy url to http request
        if (proxyURL != "")
        {
            WebProxy webProxy = new WebProxy();
            Uri proxyAddr = new Uri(proxyURL);
            webProxy.Address = proxyAddr;
            fbRequest.Proxy = webProxy;
        }

        // add add'l http request attributes
        fbRequest.ContentType = "text/xml";
        fbRequest.Method = "GET";
        fbRequest.KeepAlive = true;
        fbRequest.AllowAutoRedirect = false;

        // response will return with HTML not XML if UserAgent isn't set
        fbRequest.UserAgent = "Mozilla/4.0(compatible;MSIE 6.0; " +
            "Windows NT 5.1; SV1; .NET CLR 1.1.4322; .NET CLR 2.0.50727; InfoPath.1)";

        // return the response from http web request
        return (HttpWebResponse)fbRequest.GetResponse();
    }

    public string GetFormsAuthCookies(string mailserver, string domain, string username, string password)
    {
        byte[] postBody;
        Uri authURL = new Uri(owaURL + "/exchweb/bin/auth/owaauth.dll");
        HttpWebRequest cookieRequest;
        HttpWebResponse cookieResponse;
        Stream ioStream;
        StreamReader ioReader;
        string postString;
        string responseData;
        string cookies;
        string cadCookieData;
        string sessionIDCookie;

        // accept owa mail server x509 certificate using a callback
        ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(AcceptAllCerts);

        // create the HttpWebRequest object to the forms authentication url
        cookieRequest = (HttpWebRequest)WebRequest.Create(authURL);
        
        // if proxy is used then add proxy url to http request
        if (proxyURL != "")
        {
            WebProxy webProxy = new WebProxy();
            Uri proxyAddr = new Uri(proxyURL);
            webProxy.Address = proxyAddr;
            cookieRequest.Proxy = webProxy;
        }
        
        // create cookie container and add it to http request object
        CookieContainer CookieJar = new CookieContainer();
        cookieRequest.CookieContainer = CookieJar;

        // concatenate fields to be posted
        postString = "destination=" + owaURL + "%2Fexchange%2F" + username + "%2F&username=" + domain + "%5C" + username + "&password=" + password + "&SubmitCreds=Log+On&forcedownlevel=0&trusted=0";
        
        // add http request attributes
        cookieRequest.KeepAlive = true;
        cookieRequest.AllowAutoRedirect = false;
        cookieRequest.Method = "POST";
        cookieRequest.ContentType = "application/x-www-form-urlencoded";

        // encode string to be posted
        postBody = Encoding.UTF8.GetBytes(postString);
        cookieRequest.ContentLength = postBody.Length;

        // create the IO stream, adding the post field length and close stream
        ioStream = cookieRequest.GetRequestStream();
        ioStream.Write(postBody, 0, postBody.Length);
        ioStream.Close();

        // send the POST method request.
        cookieResponse = (System.Net.HttpWebResponse)cookieRequest.GetResponse();
        
        // read response and close response connection
        ioReader = new StreamReader(cookieResponse.GetResponseStream());
        responseData = ioReader.ReadToEnd();
        cookieResponse.Close();

        cookies = CookieJar.GetCookieHeader(authURL).ToString();

        //Filter for our cadata and session ID cookies.
        cadCookieData = Regex.Replace(cookies, "(.*)cadata=(.*)(.*)", "$2");
        sessionIDCookie = Regex.Replace(cookies, "(.*)sessionid=(.*)(,|;)(.*)", "$2");

        // Create and return the cookie set for performing subsequent Web requests.
        cookies = "sessionid=" + sessionIDCookie + "; " + "cadata=" + cadCookieData;
        return cookies;
    }
    // just accept the server certificate
    private static bool AcceptAllCerts(object sender, X509Certificate cert, X509Chain chain, SslPolicyErrors sslPolicyErrors)
    {
        return true;
    }
}

