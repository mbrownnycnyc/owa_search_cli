using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CommandLine;
using CommandLine.Text;
using System.Security;

namespace exch2007_owa_searcher
{
    class Options
    {
        //for info on mutuallyexclusivesets see:
        // http://commandline.codeplex.com/wikipage?title=MutuallyExclusiveParsingFixture_CSharp
        // http://commandline.codeplex.com/workitem/7134
        // http://stackoverflow.com/questions/10639320/command-line-parser-with-mutually-exclusive-required-parameters

        //meta:
        [Option('v', "verbose", DefaultValue = false, Required = false,
            HelpText = "Prints all messages to standard output.")]
        public bool Verbose { get; set; }

        //setup:
        [Option('u', "username", Required = true,
            HelpText = "User to log on to EWS.  Use the same credentials you use for OWA access.")]
        public string setupUsername { get; set; }

        [Option('p', "password", Required = false,
            HelpText = "At this time, there is no support for password entry from argument.  If you don't provide this at the command-line, you will be prompted at runtime (not providing at command line is more secure).  Password for user to log on to EWS provided with --username/-u.  This is stored immediately in volatile memory as a SecureString.  Use the same credentials you use for OWA access.")]
        public SecureString setupPassword { get; set; }
                
        [Option('a', "autodiscovery", DefaultValue = true, Required = false,
            HelpText = "Boolean to use if you want EWS autodiscovery to take place.")]
        public bool setupAutoDiscovery { get; set; }

        [Option('e', "email", Required = false, 
            HelpText = "Email address used for auto discovery of the EWS URL.  Use the same email you use for OWA access.")]
        public string setupEmailAddress { get; set; }

        //should never have uri, email, and autodiscovery
        //should only ever have uri + email (required), or autodiscovery
        //utilizing '\0' as Unicode null char, as ShortName is a nullable char
        [Option('l', "url", Required = false,
            HelpText = "URL used for EWS.")]
        public string setupURI { get; set; }

        
        //search related:
        [Option('f', "from", Required = false,
            HelpText = "Refer to PostItemSchema.From property.  This can be any partial hit of 'Name' or Email address.")]
        public string searchFrom { get; set; }

        [Option('t', "to", Required = false,
            HelpText = "Refer to PostItemSchema.To property.  This can be any partial hit of 'Name' or Email address.")]
        public string searchTo { get; set; }

        [Option('s', "subject", Required = false,
            HelpText = "Refer to ItemSchema.Subject property.")]
        public string searchSubject { get; set; }

        [Option('b', "body", Required = false,
            HelpText = "Refer to ItemSchema.Body property.")]
        public string searchBody { get; set; }

        [Option('w', "hasattachments", DefaultValue = false, Required = false,
            HelpText = "Boolean to return only Emails that have attachments.")]
        public bool searchHasAttachments { get; set; }

        [OptionList('i', "importance", Separator = ',',
            HelpText = "Provide one or more importance levels (low, normal, high). Refer to Microsoft.Exchange.WebServices.Data Importance enumeration.")]
        public IList<string> searchImportance { get; set; }

        [OptionList('x', "sensitivity", Separator = ',',
            HelpText = "Provide one or more sensitivity levels (normal, personal, private, confidential). Refer to Microsoft.Exchange.WebServices.Data Sensitivity enumeration.")]
        public IList<string> searchSensitivity { get; set; }
                
        //http://commandline.codeplex.com/wikipage?title=The-Option-List-Attribute&referringTitle=Documentation
        [OptionList('r', "datereceived", Separator = ';',
            HelpText = "Only available on items that you've received, specifies a range of datetimes to search date received as '[ISO 8601 start datetime];[ISO 8601 end datetime]'")]
        public IList<DateTime> searchDateReceivedRange { get; set; }
        
        [OptionList('d', "datesent", Separator = ';',
            HelpText = "Only available on items that you've Sent, specifies a range of datetimes to search date sent as '[ISO 8601 start datetime];[ISO 8601 end datetime]'")]
        public IList<DateTime> searchDateSentRange { get; set; }

        [OptionList('m', "size", Separator = ',',
            HelpText = "Providing `--size \"~N KB\"` will return all email items(i) (60%)N<i>(140%)N, aka 3KB-7KB. Acceptable units are: bytes(B), kilobytes(KB), megabytes(MB).")]
        public IList<string> searchSizeRange { get; set; }


        //search operation
        [Option('o', "logicaloperator", DefaultValue = "AND", Required = false,
            HelpText = "Boolean to return only Emails that have attachments.")]
        public string searchLogicalOperator { get; set; }

        [Option('z', "recursive", Required = false,
            HelpText = "Performs search operation recursively on given Folder.  If provided without options, Inbox is the default.")]
        public string searchRecursiveFolder { get; set; }
        
        [Option('n', "view", DefaultValue = int.MaxValue, Required = false,
            HelpText = "Creates an item view of a limited size.")]
        public int searchItemViewSize { get; set; }


        [ParserState]
        public IParserState LastParserState { get; set; }

        [HelpOption]
        public string GetUsage()
        {
            return HelpText.AutoBuild(this,
              (HelpText current) => HelpText.DefaultParsingErrorsHandler(this, current));
        }
    }
}
