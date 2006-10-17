Installing and using the DirectoryWebService example:
--------------------------------------------------------
REQUIRES: Visual Studio 2005

This service was created to provide xml data to a OneBox module for employee search	results.

ASSUMPTIONS:
		You have a Microsoft LDAP server which you can query for user attibutes such as name, phone etc.
		You have an Outlook for Web Access (OWA) server in which you can query the public mail store for
		users free/busy schedule.  If you don't have an OWA store and don't want to bring back users free/bush
		schedule just comment out line 97 of the DirectoryService.cs file.  This code was written for an environment
		where basic authentication is used over SSL (port 443) to Access	the outlook mail store and that an x509
		certificate is used.  If your mail server does not use certificates I believe the code will still work fine
		or you can just comment out line 257 of the DirectoryService.cs.

IIS SETUP / CODE MODIFICATIONS:
1)	Create a Virtual Directory on IIS called DirectoryWebSvc and point it to the location where you
		downloaded the project files.

2)	Open the DirectoryWebSvc.sln file then modify DirectoryService.cs in the App_Code directory as follows:
		2a)	line 12 - WebService Namespace - change to the URL where this code resides on your IIS
		2b)	line 19 - "Namespace" - same as above
		2c)	line 20 - "UserID" - user allowed to perform an LDAP search.  Same user in our example searches the OWA
				mail server to get users free/busy schedule
		2d) line 21 - "Password" - password for above UserID
		2e) line 22 - "Domain" - name of domain
		2f) line 23 - "MailServer" - name of Outlook for Web access (OWA) server, typically mail.mydomain.com
		2g) line 49 - Change to your LDAP server, port and search path where users are located, something
				like "LDAP://yourLDAPServer:389/ou=searchPath,dc=mydomain,dc=com"
		2h)	line 82 - Change to URL where OneBox image is located (supplied in image directory or just use your own)

USAGE:
		Compile and run the code.  You should see the URL http://localhost/DirectoryWebSvc/DirectoryService.asmx
		Click on the GoogleUsers link then enter your first and last name in the text box and click the "Invoke"
		button.  You should then see the output in XML.  At the bottom of this readme is more information about the
		anr search function in LDAP. Please read it and experiment with custom name searches.

GSA SETUP:
		Using a text editor open the attached directory_search.xml file and perform the following:
1)	line 11 - modify the variable SearchRequeryUrl to your GSA's URL
3)	line 13 - modify the href attribute for the directory.css style sheet to the URL where the directory.css
		is located.

1)	Log into the GSA as an administrator (http://myGSA.myDomain.com:8000/EnterpriseController)
2)	Select Serving > OneBox Modules from the left side of the page.  Click the "Browse..." button and import
		the provided directory_search.xml document.
3)	Change the URL in the "External Provider" text box to the URL of the DirectoryService.asmx file. If
		you're using a local IIS don't use localhost, use the name of your local computer. The GSA doesn'the
		know who localhost is. Save the changes.
4)	Now you must associate this OneBox module to a frontend.  Go to Serving and select the "OneBox Modules"
		tab then select the module from the "Available Modules" list and click the > arrow to move it to the
		"Selected Modules" list.  Save the settings.
5)	Go to the GSA search page and enter one of the keywords that trigger the onebox


ABOUT THE SEARCH:
	  This service takes 1 parameter which is an LDAP search function called anr which stands for
		ambiguous name resolution.  Here is some information about anr from http://support.microsoft.com/?kbid=243299

		LDAP clients can use ANR to make searching and querying easier. Rather than presenting complex filters,
		a search can be presented for partial matches. If a space is embedded in the search string, as in the
		case above, the search is divided at the space and an "or" search is also performed on the attributes.
		If there is more than one space, the search divides only at the first space.

		By default, the following attributes are set for ANR:
		• GivenName
		• Surname
		• displayName
		• LegacyExchangeDN
		• msExchMailNickname
		• RDN
		• physicalDeliveryOfficeName
		• proxyAddress
		• sAMAccountName
