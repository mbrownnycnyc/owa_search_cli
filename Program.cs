using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Exchange.WebServices.Data;
using System.Reflection;
using System.Threading;
using System.Text.RegularExpressions;
using CommandLine;
using CommandLine.Text;
using System.Security;

/*
todo:
 Incorporate Body preview as noted in stackoverflow thread.
 * 
 Several arguments should be parsed: (
 - user should be able to provide username and password via command line.
 - if no password provided, then prompt at runtime. use then dispose of string.
 - set URI to exchange.asmx or just use autodiscoverURL
 - set SearchFilter properties
 - set logicaloperator (default logical AND)
 - set max number of return items (default int.MaxValue)
 - search given folder and subfolders? (default)
 - search only subfolders of given folder?
 
 optionally should use an XML config file.
 - user can provide a path to this XML.
 - if no path provided, use one with the executable name in the search path or %appdata%\owa_searcher\[executablename].xml.
*/

namespace exch2007_owa_searcher
{
    class Program
    {
        //http://msdn.microsoft.com/en-us/library/ms526356%28v=EXCHG.10%29.aspx
        //http://msdn.microsoft.com/en-us/library/ms526844%28v=exchg.10%29.aspx
        //ExtendedPropertyDefinition PR_TRANSPORT_MESSAGE_HEADERS = new ExtendedPropertyDefinition(0x007D, MapiPropertyType.String);

        static void Main(string[] args)
        {
            //within main() we will perform robust checks against the options.
            //we will then call performsearch()


            //SecureString securePassword;

            //http://commandline.codeplex.com/
            //  Install-Package CommandLineParser. 
            var options = new Options();
            CommandLine.Parser.Default.ParseArguments(args, options);
            //perform validation of all arguments:

            try
            {
                if (options.setupUsername != null)
                {
                    if (options.setupPassword != null && options.setupPassword.GetType().ToString() == "System.String")
                    {
                        //SecureStringManager ssmanager = new SecureStringManager();
                        //securePassword = ssmanager.convertToSecureString(options.setupPassword);
                        //options.setupPassword = "";
                    }
                    else
                    {
                        Console.Write("password: ");
                        options.setupPassword = SecureStringManager.getPasswordCLI();
                    }
                }

                if (options.setupAutoDiscovery == true && (options.setupEmailAddress == null || !options.setupEmailAddress.Contains('@')))
                {
                    Console.Error.WriteLine("Please provide an Email address with an '@'.");
                    options.GetUsage();
                }

                if (options.setupAutoDiscovery = false & (options.setupURI == null || !options.setupURI.ToString().Contains("https://")))
                {
                    Console.Error.WriteLine("Please provide an URI that contains `https://`.");
                    options.GetUsage();
                }

                if (options.searchRecursiveFolder != null && options.searchRecursiveFolder.Contains(""))
                {
                    //do nothing
                }

                //throw new OptionException("test");

                PerformSearch(options);
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.Message);
            }

        }



        public class OptionException : Exception
        {
            public OptionException(string message): base(message){}
        }

        public static void PerformSearch(Options options)
        {

            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);

            if (options.setupAutoDiscovery)
            {
                Console.WriteLine("Attempting autodiscovery for " + options.setupEmailAddress + "...");
                service.AutodiscoverUrl(options.setupEmailAddress, RedirectionCallback);
                Console.WriteLine("Discovered: Using " + service.Url.ToString());
            }
            else if (options.setupURI != null)
            {
                service.Url = new Uri(options.setupURI);
            }

            service.UseDefaultCredentials = true;
            if (options.setupPassword.GetType() == typeof(string))
            {
                //for non-secure string
                service.Credentials = new WebCredentials(options.setupUsername, "");
            }
            else if (options.setupPassword.GetType() == typeof(SecureString))
            {
                // for securestring:
                SecureStringManager ssmanager = new SecureStringManager();
                string pword = ssmanager.convertToPlainTextString(options.setupPassword);
                service.Credentials = new WebCredentials(options.setupUsername, pword);
                pword = "";
            }

            // Add a search filter that searches on the body or subject.
            List<SearchFilter> searchFilterCollection = new List<SearchFilter>();

            // http://msdn.microsoft.com/en-us/library/microsoft.exchange.webservices.data.itemschema_fields%28v=exchg.80%29.aspx
            // http://stackoverflow.com/a/15558850/843000 + http://stackoverflow.com/a/2551057/843000



            if (options.searchBody != null)
            {
                searchFilterCollection.Add(new SearchFilter.ContainsSubstring(PostItemSchema.Body, options.searchBody));
            }
            if (options.searchDateReceivedRange != null)
            {
                //searchFilterCollection.Add(new SearchFilter.ContainsSubstring(PostItemSchema.From, options.searchFrom));
                Console.WriteLine("datercvd not enabled yet");
            }
            if (options.searchDateSentRange != null)
            {
                //searchFilterCollection.Add(new SearchFilter.ContainsSubstring(PostItemSchema.From, options.searchFrom));
                Console.WriteLine("datesent not enabled yet");
            }
            if (options.searchFrom != null)
            {
                searchFilterCollection.Add(new SearchFilter.ContainsSubstring(PostItemSchema.From, options.searchFrom));
            }
            if (options.searchHasAttachments == true)
            {
                //searchFilterCollection.Add(new SearchFilter.ContainsSubstring(PostItemSchema.HasAttachments, options.searchHasAttachments));
                Console.WriteLine("hasattachment not enabled yet");
            }
            if (options.searchImportance != null)
            {
                //searchFilterCollection.Add(new SearchFilter.ContainsSubstring(PostItemSchema.From, options.searchFrom));
                Console.WriteLine("importance not enabled yet");
            }
            if (options.searchSensitivity != null)
            {
                //searchFilterCollection.Add(new SearchFilter.ContainsSubstring(PostItemSchema.From, options.searchFrom));
                Console.WriteLine("sensitivity not enabled yet");
            }
            if (options.searchSizeRange != null)
            {
                //searchFilterCollection.Add(new SearchFilter.ContainsSubstring(PostItemSchema.From, options.searchFrom));
                Console.WriteLine("sizerange not enabled yet");
            }
            if (options.searchSubject != null)
            {
                searchFilterCollection.Add(new SearchFilter.ContainsSubstring(PostItemSchema.Subject, options.searchSubject));
            }
            if (options.searchTo != null)
            {
                searchFilterCollection.Add(new SearchFilter.ContainsSubstring(PostItemSchema.Sender, options.searchTo));
            }



            //searchFilterCollection.Add(new SearchFilter.ContainsSubstring(ItemSchema.Subject, "Out next week for two weeks"));
            //searchFilterCollection.Add(new SearchFilter.ContainsSubstring(ItemSchema.Body, "homecoming"));
            //searchFilterCollection.Add(new SearchFilter.IsGreaterThan(ItemSchema.DateTimeReceived, DateTime.Now.AddHours(24)));


            // Create the search filter.
            //if (options.searchLogicalOperator.ToLower() == "and")
            //{
             SearchFilter searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, searchFilterCollection.ToArray());
            //}
            if (options.searchLogicalOperator.ToLower() == "or")
            {
                searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.Or, searchFilterCollection.ToArray());
            }


            // Create a view with a page size of 50.
            //ItemView itemView = new ItemView(50);
            ItemView itemView = new ItemView(int.MaxValue);

            // Identify the Subject and DateTimeReceived properties to return.
            // Indicate that the base property will be the item identifier
            itemView.PropertySet = new PropertySet(BasePropertySet.IdOnly,
                ItemSchema.Subject,
                ItemSchema.DateTimeReceived,
                PostItemSchema.Sender,
                PostItemSchema.From,
                ItemSchema.Id
                );

            // Order the search results by the DateTimeReceived in descending order.
            itemView.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Descending);

            // Set the traversal to shallow. (Shallow is the default option; other options are Associated and SoftDeleted.)
            itemView.Traversal = ItemTraversal.Shallow;


            // Send the request to search the Inbox and get the results.
            //http://stackoverflow.com/a/15276675/843000
            List<Folder> completeListOfFolderIds = new List<Folder>();

            //Fills list with all folders and subfolders
            GetAllSubFolders(service, WellKnownFolderName.Inbox, completeListOfFolderIds);


            Console.WriteLine("Searching root");
            FindItemsResults<Item> TargetRootFolder = service.FindItems(WellKnownFolderName.Inbox, searchFilter, itemView);

            //should break off into a function
            Regex serviceURL = new Regex(@"https://([\w+?\.\w+])+([a-zA-Z0-9\~\!\@\#\$\%\^\&\*\(\)_\-\=\+\\\?\.\:\;\'\,]*)?/");

            foreach (Folder folder in completeListOfFolderIds)
            {
                Console.WriteLine("Searching... " + folder.DisplayName);
                //http://stackoverflow.com/a/590999/843000

                FindItemsResults<Item> searchResults = service.FindItems(folder.Id, searchFilter, itemView);

                IEnumerable<Item> totalresults = searchResults.Union(TargetRootFolder);

                //do something with item list
                // Process each item.

                /*
                Console.WriteLine("--------------------------------------------------------------------------------");
                Console.Write("|");
                Console.Write("datetime rcvd".PadRight(19));
                Console.Write("|");
                */

                //foreach (Item myItem in totalresults)
                foreach (object myItem in totalresults)
                {
                    if (myItem is EmailMessage)
                    {
                        // http://msdn.microsoft.com/en-us/library/exchange/microsoft.exchange.data.transport.email.emailmessage(v=exchg.150).aspx


                        // for body: http://stackoverflow.com/a/14881281/843000
                        // this is also interesting: http://www.infinitec.de/post/2009/06/09/Getting-the-body-of-an-Email-with-a-FindItems-request.aspx

                        //should convert DateTimeReceived to 'YYYY-MM-dd HH:mm:ss'
                        Console.WriteLine("");
                        Console.WriteLine((myItem as EmailMessage).DateTimeReceived);
                        Console.WriteLine("----");
                        Console.WriteLine((myItem as EmailMessage).From.Name);
                        Console.WriteLine("----");
                        Console.WriteLine((myItem as EmailMessage).Subject);
                        Console.WriteLine("----");

                        /*
                        bool getbody = true;

                        if (getbody) {
                            PropertySet propset = new PropertySet(BasePropertySet.FirstClassProperties, EmailMessageSchema.From, EmailMessageSchema.ToRecipients);
                            EmailMessage.Bind(service, (myItem as EmailMessage).Id, propset);
                            Console.WriteLine((myItem as EmailMessage).Body);
                            Console.WriteLine("----");
                        }
                        Console.WriteLine(
                            (myItem as Item).Body.ToString()

                            //(myItem as EmailMessage).Body.Text.Substring(0,50)
                            
                            );
                        Console.WriteLine("----");
                         */

                        // regex handles only the domain name.
                        String sEntryID = (myItem as EmailMessage).Id.ToString();
                        String sEWSID = GetConvertedEWSID(service, sEntryID, options.setupEmailAddress);

                        Console.WriteLine(serviceURL.Match(service.Url.ToString()).Value + @"owa/?ae=Item&t=IPM.Note&id=" + sEWSID);
                    }

                    else if (myItem is MeetingRequest)
                    {
                        Console.WriteLine((myItem as MeetingRequest).DateTimeReceived);
                        Console.WriteLine((myItem as MeetingRequest).From);
                        Console.WriteLine((myItem as MeetingRequest).Subject);
                    }
                    else
                    {
                        // Else handle other item types.
                    }
                }

            }
            //end: http://stackoverflow.com/a/15276675/843000
        }

        //http://msdn.microsoft.com/en-us/library/exchange/dd635285%28v=exchg.80%29.aspx
        static bool RedirectionCallback(string url)
        {
            // Return true if the URL is an HTTPS URL.
            bool validateHttps = url.ToLower().StartsWith("https://");
            return validateHttps;
        }

        //http://blogs.msdn.com/b/brijs/archive/2010/09/09/how-to-convert-exchange-item-s-entryid-to-ews-unique-itemid-via-ews-managed-api-convertid-call.aspx
        //http://gsexdev.blogspot.com/2010_05_01_archive.html
        public static String GetConvertedEWSID(ExchangeService esb, String sID, String strSMTPAdd)
        {
            // Create a request to convert identifiers.
            AlternateId objAltID = new AlternateId();
            objAltID.Format = IdFormat.EwsId;
            objAltID.Mailbox = strSMTPAdd;
            objAltID.UniqueId = sID;

            AlternateIdBase objAltIDBase = esb.ConvertId(objAltID, IdFormat.OwaId);
            AlternateId objAltIDResp = (AlternateId)objAltIDBase;
            return objAltIDResp.UniqueId;
        }

        //http://stackoverflow.com/a/15276675/843000
        static void GetAllSubFolders(ExchangeService service, FolderId searchTargetParentFolderID, List<Folder> completeListOfFolderIds)
        {

            FolderView folderView = new FolderView(int.MaxValue);

            FindFoldersResults findFolderResults = service.FindFolders(searchTargetParentFolderID, folderView);

            List<string> foldertargetlist = new List<string>();
            foreach (Folder folder in findFolderResults)
            {
                completeListOfFolderIds.Add(folder);
                foldertargetlist.Add(folder.DisplayName);
                FindAllSubFolders(service, folder.Id, completeListOfFolderIds);
            }

            Console.WriteLine("Located folders: " + string.Join(", ", foldertargetlist.ToArray()));
        }

        static private void FindAllSubFolders(ExchangeService service, FolderId parentFolderId, List<Folder> completeListOfFolderIds)
        {
            //search for sub folders
            FolderView folderView = new FolderView(int.MaxValue);
            FindFoldersResults foundFolders = service.FindFolders(parentFolderId, folderView);

            // Add the list to the growing complete list
            completeListOfFolderIds.AddRange(foundFolders);

            // Now recurse
            foreach (Folder folder in foundFolders)
            {
                FindAllSubFolders(service, folder.Id, completeListOfFolderIds);
            }
        }

    }
}
