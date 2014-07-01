using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.DirectoryServices;
using System.IO;
using Microsoft.Web.Administration;

namespace EnableMasterPassword
{
    class Program
    {

        static String masterPassword =  "IZ8jz3j9pQz831UH";
        private static String administratorUserName = "Administrator";
        private static String administratorPassword = "CynErgi9";
        private static String ldapPath = "LDAP://iis-ad.iis.netregistry.net/ou=clients,dc=iis,dc=netregistry,dc=net";


        static void Main(string[] args)
        {
           // changeToMasterPasswordForIIS7Users();

            //Console.WriteLine(findSAMNameByDomain("joelin.au.c1om"));

            changeToMasterPasswordForIIS7Users();

        }


        static void fixOldIIS7Name()
        {
            ServerManager serverManager = new ServerManager();
            var ldap = new DirectoryEntry(ldapPath, administratorUserName, administratorPassword);


            var searcher = new DirectorySearcher(ldap);

            searcher.PropertiesToLoad.Add("displayName");
            SearchResultCollection results = searcher.FindAll();

            for (int i = 0; i < results.Count; i++)
            {
                SearchResult result = results[i];
                DirectoryEntry user = result.GetDirectoryEntry();
                String description = (string) user.Properties["Description"].Value;


                if (description == null ||
                    (!description.EndsWith(" iis7-1") && !description.EndsWith(" iis6") && !description.EndsWith(" fp")))
                {



                    string account = (string) user.Properties["SAMAccountName"].Value;

                    Console.WriteLine("old account name = " + account);

                    string domainName = null;

                    if (account != null && account.EndsWith("_iis7") && account.StartsWith("jojo1427"))
                    {

                        Console.WriteLine("find it...." + account);

                        //1) rename 
                        account = account.Replace("_iis7", "");


                        Console.WriteLine("new account name = " + account);

                        user.Rename("cn=" + account);

                        user.Properties["SAMAccountName"].Value = account;


                        domainName = ((string) user.Properties["description"].Value).Split(' ')[0];





                        //2) change password
                        user.Invoke("SetPassword", new object[] {masterPassword}); // user.userPassword 


                        //3) change description 
                        user.Properties["Description"].Value = domainName + " iis7-1";


                        //4) commit all changes
                        user.CommitChanges();
                        ldap.CommitChanges();

                        //4) update website;


                        serverManager.ApplicationPools[domainName].ProcessModel.Password = masterPassword;


                        Site site = serverManager.Sites[domainName];


                        for (int j = 0; j < site.Applications.Count; j++)
                        {
                            Console.WriteLine("there are  " + site.Applications.Count + " applications");

                            for (int k = 0; k < site.Applications[j].VirtualDirectories.Count; k++)
                            {
                                Console.WriteLine("[" + j + "]and there are " +
                                                  site.Applications[j].VirtualDirectories.Count + " virtual directories");

                                Console.WriteLine("Application Path Name = " + site.Applications[j].Path);




                                Console.WriteLine("Physical Path = " +
                                                  site.Applications[j].VirtualDirectories[k].PhysicalPath);
                                Console.WriteLine("Username = " + "IIS\\" + user.Properties["SAMAccountName"].Value);
                                Console.WriteLine("Password = " + masterPassword);

                                site.Applications[j].ApplicationPoolName = domainName;




                                site.Applications[j].VirtualDirectories[k].LogonMethod =
                                    AuthenticationLogonMethod.ClearText;

                                site.Applications[j].VirtualDirectories[k].UserName = "IIS\\" +
                                                                                      user.Properties["SAMAccountName"].
                                                                                          Value;
                                site.Applications[j].VirtualDirectories[k].Password = masterPassword;

                            }
                        }
                    }
                    serverManager.CommitChanges();
                }
            }
        }










        private static String findSAMNameByDomain(string domain)
        {
            var ldap = new DirectoryEntry(ldapPath, administratorUserName, administratorPassword);
            var searcher = new DirectorySearcher(ldap);
            searcher.PropertiesToLoad.Add("displayName");
            searcher.Filter = ("Description=" + domain + " iis7-1");

            SearchResultCollection results = searcher.FindAll();

            for (int i = 0; i < results.Count; i++)
            {
                SearchResult result = results[i];
                DirectoryEntry user = result.GetDirectoryEntry();

                return (string)user.Properties["SAMAccountName"].Value;

            }
            return null;
        }



        static void changeToMasterPasswordForIIS7Users()
        {
            

            ServerManager serverManager = new ServerManager();
            var ldap = new DirectoryEntry(ldapPath, administratorUserName, administratorPassword);


            var searcher = new DirectorySearcher(ldap);

            searcher.PropertiesToLoad.Add("displayName");
            SearchResultCollection results = searcher.FindAll();

            for (int i = 0; i < results.Count; i++)
            {
                SearchResult result = results[i];
                DirectoryEntry user = result.GetDirectoryEntry();


                if (user.Properties["Description"].Value != null &&
                    ((string) user.Properties["Description"].Value).EndsWith("iis7-1")                    
                   )
                {
                    // && ((string)user.Properties["Description"].Value).StartsWith("tildablasta.au.com")
                    Console.WriteLine("Found user" + user.Properties["SAMAccountName"]);

                    String description = (string) user.Properties["Description"].Value;
                    user.Invoke("SetPassword", new object[] { masterPassword }); // user.userPassword 

                    user.CommitChanges();
                    ldap.CommitChanges();

                    Console.WriteLine(user.Properties["Description"].Value);

                   
                    string domain = description.Split(' ')[0];

                    Console.WriteLine("domain is " + domain);

                    serverManager.ApplicationPools[domain].ProcessModel.Password = masterPassword;


                    Site site = serverManager.Sites[domain];

                    
                    for(int j=0; j<site.Applications.Count; j++)
                    {
                        Console.WriteLine("there are  " + site.Applications.Count + " applications");

                        for(int k=0;k< site.Applications[j].VirtualDirectories.Count; k++)
                        {
                            Console.WriteLine("[" +  j +  "]and there are " + site.Applications[j].VirtualDirectories.Count + " virtual directories");

                            Console.WriteLine("Application Path Name = " + site.Applications[j].Path);

                            


                            Console.WriteLine("Physical Path = " + site.Applications[j].VirtualDirectories[k].PhysicalPath);
                            Console.WriteLine("Username = " + "IIS\\" +user.Properties["SAMAccountName"].Value);
                            Console.WriteLine("Password = " + masterPassword);

                            site.Applications[j].ApplicationPoolName = domain;


                            

                            site.Applications[j].VirtualDirectories[k].LogonMethod = AuthenticationLogonMethod.ClearText;

                            site.Applications[j].VirtualDirectories[k].UserName = "IIS\\" +
                                                                                  user.Properties["SAMAccountName"].Value;
                            site.Applications[j].VirtualDirectories[k].Password = masterPassword;
                           
                        }
                    }
                }
                serverManager.CommitChanges();
            }
        }
    }
}
