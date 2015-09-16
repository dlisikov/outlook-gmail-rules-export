using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.Outlook;

namespace OutlookRulesExport
{
    class Program
    {
        static void Main(string[] args)
        {
            Store store;
            String storeName;

            if (args.Count() == 0)
            {
                PrintHelp();
                return;
            }

            // print the stores(mailboxes) available
            if (args[0] == "-s")
            {
                GetStores();
                return;
            }

            // print out rules in CSV format
            if (args[0] == "-c")
            {
                if (args.Count() >= 2)
                {
                    storeName = args[1];

                    store = GetStore(storeName);
                    if (store != null)
                    {
                        List<ParsedRule> rulesList = GetRules(store, storeName);
                        PrintCSV(rulesList);
                    }
                }

                return;
            }

            // default, print rules to xml file
            storeName = args[0];
            store = GetStore(storeName);
            if (store != null)
            {
                var rulesList = GetRules(store, storeName);
                File.WriteAllLines("rules.txt", rulesList.Select(rule => rule.ToString()));
            }
        }

        private static void PrintHelp()
        {
            Console.WriteLine(String.Empty);
            Console.WriteLine("Exports rules from oulook so they can be directly imported into Gmail as filters to a rules.xml file in the current directory. No editing required.");
            Console.WriteLine(String.Empty);
            Console.WriteLine("Usage: OutlookRulesExport.exe [options] [mailbox]");
            Console.WriteLine(String.Empty);
            Console.WriteLine("Example: OutlookRulesExport.exe example@example.com");
            Console.WriteLine("Example: OutlookRulesExport.exe -c example@example.com");
            Console.WriteLine("Example: OutlookRulesExport.exe -s");
            Console.WriteLine(String.Empty);
            Console.WriteLine("-s: returns a list of stores (mailboexes) available");
            Console.WriteLine("-c: prints rules to the console in CSV format (unedited)");
        }

        /// <summary>
        /// Returns a list of stores (mailboxes and pst files) available in outlook
        /// Some or all of these stores can contain rules that can be exported
        /// </summary>
        public static void GetStores()
        {
            var app = new Application();
            Stores stores = app.Session.Stores;

            foreach (Store s in stores)
            {
                Console.WriteLine(s.DisplayName);
            }
        }

        /// <summary>
        /// Gets the Outlook store (mailbox) for the given string if one exists.
        /// </summary>
        /// <param name="storeName"></param>
        /// <returns>Null if no store is found</returns>
        public static Store GetStore(string storeName)
        {
            var app = new Application();
            Stores stores = app.Session.Stores;

            if (stores.Count > 0)
            {
                try
                {
                    Store s = stores[storeName];

                    if (s != null)
                    {
                        return s;
                    }
                    else
                    {
                        Console.WriteLine("Invalid mailbox");
                        PrintHelp();
                    }
                }
                catch (System.Runtime.InteropServices.COMException e)
                {
                    Console.WriteLine("No such mailbox");
                    PrintHelp();
                }
            }
            else
            {
                Console.WriteLine("No mailboxes in Outlook");
                PrintHelp();
            }

            return null;
        }

        /// <summary>
        /// Gets a list of rules and associated actions from the given store
        /// Currently only working for rules defined on from addresses, will add support for other rules soon
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static List<ParsedRule> GetRules(Store s, string storeName)
        {
            Rules rules = s.GetRules();
            List<ParsedRule> rulesList = new List<ParsedRule>();

            foreach (Rule r in rules)
            {
                if (r.Enabled)
                {
                    ParsedRule mr = new ParsedRule();

                    ParseFromAddresses(r, mr);
                    ParseToAddresses(r, mr);
                    ParseLabelMove(r, mr, storeName);
                    ParseLabelCopy(r, mr, storeName);
                    ParseSubject(r, mr);
                    ParseBody(r, mr);
                    ParseBodyOrSubject(r, mr);
                    ParseMoveToTrash(r, mr);
                    ParseDeletePermamently(r, mr);
                    ParseCC(r, mr);
                    ParseToOrCC(r, mr);
                    ParseWordsInSenderAddressses(r, mr);
                    ParseWordsInRecipientAddressses(r, mr);

                    rulesList.Add(mr);
                }
            }

            return rulesList;
        }

        private static void ParseBody(Rule r, ParsedRule mr)
        {
            if (r.Conditions.Body.Enabled)
            {
                string[] temp = r.Conditions.Body.Text;

                for (int i = 0; i < temp.Length; i++)
                {
                    mr.BodyContains += temp[i];

                    if (i != temp.Length - 1)
                        mr.BodyContains += " OR ";
                }
            }
        }
        
        private static void ParseBodyOrSubject(Rule r, ParsedRule mr)
        {
            if (r.Conditions.BodyOrSubject.Enabled)
            {
                string[] temp = r.Conditions.BodyOrSubject.Text;

                for (int i = 0; i < temp.Length; i++)
                {
                    mr.BodyOrSubjectContains += temp[i];

                    if (i != temp.Length - 1)
                        mr.BodyOrSubjectContains += " OR ";
                }
            }
        }

        private static void ParseSubject(Rule r, ParsedRule mr)
        {
            if (r.Conditions.Subject.Enabled)
            {
                string[] temp = r.Conditions.Subject.Text;

                for (int i = 0; i < temp.Length; i++)
                {
                    mr.SubjectContains += temp[i];

                    if (i != temp.Length - 1)
                        mr.SubjectContains += " OR ";
                }
            }
        }

        private static void ParseLabelMove(Rule r, ParsedRule mr, string storeName)
        {
            try
            {
                if (r.Actions[1].ActionType == OlRuleActionType.olRuleActionMoveToFolder)
                {
                    if (r.Actions.MoveToFolder.Enabled)
                    {
                        MAPIFolder folder = r.Actions.MoveToFolder.Folder;
                        if (folder != null)
                        {
                            mr.MoveToFolder = CleanRuleActions(folder.FolderPath, storeName);
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        private static void ParseMoveToTrash(Rule r, ParsedRule mr)
        {
            try
            {
                if (r.Actions.Delete.Enabled)
                {
                    mr.MoveToTrash = true;
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        private static void ParseDeletePermamently(Rule r, ParsedRule mr)
        {
            try
            {
                if (r.Actions.DeletePermanently.Enabled)
                {
                    mr.DeletePermanently = true;
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        private static void ParseLabelCopy(Rule r, ParsedRule mr, string storeName)
        {
            try
            {
                if (r.Actions.CopyToFolder.Enabled)
                {
                    MAPIFolder folder = r.Actions.CopyToFolder.Folder;
                    if (folder != null)
                    {
                        mr.CopyToFolder = CleanRuleActions(folder.FolderPath, storeName);
                    }
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        private static void ParseWordsInSenderAddressses(Rule r, ParsedRule mr)
        {
            if (r.Conditions.SenderAddress.Enabled &&
                    r.Conditions.SenderAddress.Address.Length > 0)
            {
                mr.SenderAddressContains = string.Join(" OR ", r.Conditions.SenderAddress.Address);
            }
        }

        private static void ParseWordsInRecipientAddressses(Rule r, ParsedRule mr)
        {
            if (r.Conditions.RecipientAddress.Enabled &&
                    r.Conditions.RecipientAddress.Address.Length > 0)
            {
                mr.RecipientAddressContains = string.Join(" OR ", r.Conditions.RecipientAddress.Address);
            }
        }

        private static void ParseFromAddresses(Rule r, ParsedRule mr)
        {
            if (r.Conditions.From.Recipients.Count > 0)
            {
                for (int i = 1; i <= r.Conditions.From.Recipients.Count; i++)
                {
                    string temp = "";
                    // voodo to extract email addresses
                    try
                    {
                        OlAddressEntryUserType addressType = r.Conditions.From.Recipients[i].AddressEntry.AddressEntryUserType;

                        if ((addressType == OlAddressEntryUserType.olExchangeRemoteUserAddressEntry) || (addressType == OlAddressEntryUserType.olExchangeUserAddressEntry))
                        {
                            temp = r.Conditions.From.Recipients[i].AddressEntry.GetExchangeUser().PrimarySmtpAddress;
                        }
                        else
                        {
                            if (addressType == OlAddressEntryUserType.olSmtpAddressEntry)
                            {
                                temp = r.Conditions.From.Recipients[i].AddressEntry.Address;
                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        Console.WriteLine(ex);
                    }

                    // compose the address string if there are mutlitple addresses in the from
                    if (!String.IsNullOrEmpty(temp))
                    {
                        if (i == 1)
                        {
                            mr.FromAddress += temp;
                        }
                        else
                        {
                            mr.FromAddress += "," + temp;
                        }
                    }
                }
            }
        }
        
        private static void ParseCC(Rule r, ParsedRule mr)
        {
            if (r.Conditions.CC.Enabled)
            {
                mr.MeInCC = true;
            }
        }

        private static void ParseToOrCC(Rule r, ParsedRule mr)
        {
            if (r.Conditions.ToOrCc.Enabled)
            {
                mr.MeDirectlyOrInCC = true;
            }
        }
        
        private static void ParseToAddresses(Rule r, ParsedRule mr)
        {
            if (r.Conditions.SentTo.Recipients.Count > 0)
            {
                for (int i = 1; i <= r.Conditions.SentTo.Recipients.Count; i++)
                {
                    string temp = "";
                    // voodo to extract email addresses
                    try
                    {
                        OlAddressEntryUserType addressType = r.Conditions.SentTo.Recipients[i].AddressEntry.AddressEntryUserType;

                        if ((addressType == OlAddressEntryUserType.olExchangeRemoteUserAddressEntry) || (addressType == OlAddressEntryUserType.olExchangeUserAddressEntry))
                        {
                            temp = r.Conditions.SentTo.Recipients[i].AddressEntry.GetExchangeUser().PrimarySmtpAddress;
                        }
                        else
                        {
                            if (addressType == OlAddressEntryUserType.olSmtpAddressEntry)
                            {
                                temp = r.Conditions.SentTo.Recipients[i].AddressEntry.Address;
                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        Console.WriteLine(ex);
                    }

                    // compose the address string if there are mutlitple addresses in the sentTo
                    if (!String.IsNullOrEmpty(temp))
                    {
                        if (i == 1)
                        {
                            mr.ToAddress += temp;
                        }
                        else
                        {
                            mr.ToAddress += "," + temp;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Print the rules in CSV format to the console
        /// </summary>
        /// <param name="rules"></param>
        public static void PrintCSV(List<ParsedRule> rules)
        {
            // print the rules
            foreach (ParsedRule mr in rules)
            {
                Console.WriteLine(mr.FromAddress + ";" + mr.MoveToFolder);
            }
        }

        /// <summary>
        /// Cleans the actions in the rule list to remove the storename and formats for 
        /// Google.
        /// </summary>
        /// <param name="rules"></param>
        /// <param name="storeName"></param>
        public static string CleanRuleActions(string path, String storeName)
        {
            // remove the store name from the action path
            path =  path.Replace("\\\\"+storeName+"\\", String.Empty);

            // swap backslash for forwardslash to play nice with google
            path = path.Replace("\\", "/");

            return path;
        }
    }
}
