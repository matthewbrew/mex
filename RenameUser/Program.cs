using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.IO;


namespace RenameUser
{



    class Program
    {

        public static string LDAP_USER_PATH = "LDAP://iis-ad.iis.netregistry.net/ou=clients,dc=iis,dc=netregistry,dc=net";
        private readonly char[] ADMINISTRATOR_PASSWORD = { 'C', 'y', 'n', 'E', 'r', 'g', 'i', '9' };
        private readonly string ADMINISTRATOR_USER_NAME = "Administrator";


        static void Main(string[] args)
        {

//            var ldap = new DirectoryEntry(LDAP_USER_PATH, "Administrator", "CynErgi9");
//
//            var searcher = new DirectorySearcher(ldap);
//            //            searcher.Filter = "(cn=coco2223)";
//            //            SearchResultCollection results = searcher.FindAll();
//            //
//            //            Console.WriteLine("results.Count=" + results.Count);
//
//            searcher.Filter = "(SAMAccountName=coco2223)";
//            searcher.PropertiesToLoad.Add("displayName");
//
//            SearchResult result = searcher.FindOne();
//
//            if (result == null)
//            {
//                Console.WriteLine("result is null");
//            }
//            else
//            {
//                Console.WriteLine("result=" + result);
//
//                Console.WriteLine("entry=" + result.GetDirectoryEntry());
//            }

            changePassword("c:\\list.txt");         
        }




        //NOT USING in IIS7
        public static void changePassword(string filename)
        {
            var ldap = new DirectoryEntry(LDAP_USER_PATH, "Administrator", "CynErgi9");
            try
            {
                var sr = new StreamReader(filename);
                string data = "";
                while ((data = sr.ReadLine()) != null)
                {
                    string username;
                    string type;

                    try
                    {
                        String[] parts = data.Split('|');

                        username = parts[0].Trim();
                        type = parts[1].Trim();

                       // Console.WriteLine("Ready to update " + username + " for hosting type " + type);
                        var searcher = new DirectorySearcher(ldap);
                        //            searcher.Filter = "(cn=coco2223)";
                        //            SearchResultCollection results = searcher.FindAll();
                        //
                        //            Console.WriteLine("results.Count=" + results.Count);

                        searcher.Filter = "(SAMAccountName="   +  username  +     ")";
                        searcher.PropertiesToLoad.Add("displayName");

                        SearchResult result = searcher.FindOne();

                        if (result != null)
                        {
                            DirectoryEntry entry = result.GetDirectoryEntry();
                            String description =  (String)  entry.Properties["Description"].Value;                           
                            
                            if(description!=null)
                            {
                                String[] coms = description.Split(' ');
                                //doing changes
                                if (type.Equals("windows"))
                                {
                                    if (!description.EndsWith(" iis6"))
                                    {
                                        //description =
                                        entry.Properties["Description"].Value = coms[0] + " iis6";
                                        entry.CommitChanges();
                                    }

                                }
                                else if (type.Equals("windows2008-1"))
                                {
                                    if (!description.EndsWith(" iis7-1"))
                                    {
                                        entry.Properties["Description"].Value = coms[0] + " iis7-1";
                                        entry.CommitChanges();
                                    }
                                }
                                else if (type.Equals("windows_fp"))
                                {
                                    if (!description.EndsWith(" fp"))
                                    {
                                        entry.Properties["Description"].Value = coms[0] + " fp";
                                        entry.CommitChanges();
                                    }

                                }

                             //   Console.WriteLine("" + username + " " + description);
                                
                            }
                            else
                            {
                                Console.WriteLine("What!?! user " + username + " does not have description");
                            }
                            
                        }
                        else
                        {
                         //   Console.WriteLine("result=" + result);
                            //Console.WriteLine("entry=" + result.GetDirectoryEntry());
                        }


                    }
                    catch (Exception e)
                    {
                        Console.WriteLine( " failed change password: " + e.Message);
                        Console.WriteLine(e.StackTrace);
                    }
                }
                sr.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
            }
        }
    }
}
