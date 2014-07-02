using System;
using System.Collections;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using log4net;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using System.Text;
using System.Management.Automation;
using System.Management.Automation.Host;
using System.Management.Automation.Runspaces;
using System.Security;

namespace Exchange
{
    public class LocalExchange : IDisposable
    {
        private static string uri = "https://mail.uc-lab.server-secure.com/powershell";
        private static string schemaURI = "http://schemas.microsoft.com/powershell/testShell";
        private static string user = "GFL\\vlad";
        private static string password = "Password1";
        private static Runspace _runspace = null;
        private static object _runspaceLock = new object();
        private static object _runspaceStateLock = new object();

        private static readonly ILog log = LogManager.GetLogger(typeof(LocalExchange));

        private static Runspace GetRunspace()
        {

            //ensure runspace exists
            if (_runspace == null)
            {
                lock (_runspaceLock)
                {

                    if (_runspace == null)
                    {
                        InitialSessionState initial = InitialSessionState.CreateDefault2();

                        //initial.ImportPSModule(new string[] { "ActiveDirectory", "DETManagement" });

                        //string pwd = "Password1";
                        //char[] cpwd = pwd.ToCharArray();
                        //System.Security.SecureString ss = new System.Security.SecureString();
                        //foreach (char c in cpwd)
                        //    ss.AppendChar(c);                        

                        //PSCredential cred = new PSCredential("ProvisioningAccount", ss);

                        //WSManConnectionInfo connectionInfo = new WSManConnectionInfo(new Uri("http://eq001yex.CIDPOC.Server-Complex.com/PowerShell/"), "http://schemas.microsoft.com/powershell/Microsoft.Exchange",
                        //    cred);


                        //connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Kerberos;
                        //connectionInfo.NoEncryption = true;


                        _runspace = RunspaceFactory.CreateRunspace(initial);


                    }
                }
            }

            //ensure runspace is still open
            if (_runspace.RunspaceStateInfo.State != RunspaceState.Opened)
            {
                lock (_runspaceStateLock)
                {
                    if (_runspace.RunspaceStateInfo.State != RunspaceState.Opened)
                    {
                        try
                        {
                            _runspace.Open();
                            RunspaceInvoke invoker = new RunspaceInvoke(_runspace);
                            invoker.Invoke("Set-ExecutionPolicy Bypass");
                            Pipeline pipeline = _runspace.CreatePipeline();
    //                        pipeline.Commands.AddScript("$secpasswd = ConvertTo-SecureString \"" + password + "\" -AsPlainText -Force");
    //                        pipeline.Commands.AddScript("$mycreds = New-Object System.Management.Automation.PSCredential (\"" + user + "\", $secpasswd)");
    //                        pipeline.Commands.AddScript("$s = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri " + uri + " -Authentication Basic -Credential $mycreds");
     //                       pipeline.Commands.AddScript("Import-PSSession $s");
                            string importModule = "Import-Module C:\\provisioning\\testModule.psm1";
                            pipeline.Commands.AddScript(importModule);
                            //pipeline.Commands.AddScript("Invoke-Command -Session $s -ScriptBlock {Import-Module testmodule}");
                            //pipeline.Commands.AddScript("new-MEXSession");
                            //pipeline.Commands.AddScript("new-MEXSession");
                            pipeline.Invoke();
                            log.Info("added test module script: " + importModule);
                            string errors = GetErrors(pipeline);
                            log.Error(errors);
                        }
                        catch(Exception e)
                        {
                            log.Error("Exception in startup", e);
                        }
                    }
                }
            }

            return _runspace;
        }

        public static void CreateRunspace()
        {
            var runspace = GetRunspace();
        }

        public static void CloseRunspace()
        {
            var runspace = GetRunspace();

            runspace.Close();
            runspace.Dispose();

        }

        public string GetMailbox()
        {
            var runspace = GetRunspace();

            // create a pipeline and feed it the script text






            //var ps = System.Management.Automation.PowerShell.Create();
            //ps.Commands.AddCommand("Import-Module").AddParameter("ActiveDirectory");
            //ps.Invoke();

            //PSCommand cmd = ps.Commands.AddCommand("Get-Process");


            //            Collection<PSObject> results = ps.Invoke();


            //// convert the script result into a single string

            //StringBuilder stringBuilder = new StringBuilder();
            //foreach (PSObject obj in results)
            //{
            //    stringBuilder.AppendLine(obj.ToString());

            //}

            Pipeline pipeline = runspace.CreatePipeline();
            pipeline.Commands.AddScript("Get-Mailbox");
            //pipeline.Commands.AddScript("Get-Module");
            //pipeline.Commands.AddScript(@"'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1'", false);
            //pipeline.Commands.AddScript("Get-Process");
            //pipeline.Commands.AddScript("Connect-ExchangeServer -auto -ClientApplication:ManagementShell");
            //pipeline.Commands.AddScript("[Environment]::MachineName");            



            // add an extra command to transform the script
            // output objects into nicely formatted strings

            // remove this line to get the actual objects
            // that the script returns. For example, the script

            // "Get-Process" returns a collection
            // of System.Diagnostics.Process instances.

            //pipeline.Commands.Add("Out-String");

            // execute the script

            var results = pipeline.Invoke();


            // convert the script result into a single string

            var stringBuilder = new StringBuilder();
            foreach (PSObject obj in results)
            {
                stringBuilder.AppendLine(obj.ToString());
            }

            return stringBuilder.ToString();
        }

        public string GetMtMailbox()
        {
            var runspace = GetRunspace();
            Pipeline pipeline = runspace.CreatePipeline();
            pipeline.Commands.AddScript("Get-MTMailbox -Identity mb000013a");

            try
            {
                var results = pipeline.Invoke();
                string errors = GetErrors(pipeline);
                if (errors != "")
                {
                    //fail !
                }

                // convert the script result into a single string

                var stringBuilder = new StringBuilder();
                foreach (PSObject obj in results)
                {
                    stringBuilder.AppendLine(obj.ToString() + "\n");
                    foreach (PSPropertyInfo info in obj.Properties)
                    {
                        stringBuilder.AppendLine(info.ToString() + "\n");
                    }
                    string forwardingAddress = GetString(obj, "ForwardingAddress");
                    stringBuilder.AppendLine("Forwarding address: " + forwardingAddress + "\n");

                    foreach (Object alias in GetObjectArray(obj, "EmailAlias"))
                    {
                        stringBuilder.AppendLine("Email Aliases: " + alias); 
                    }
                }

                return stringBuilder.ToString();
            }
            catch(Exception e)
            {
                log.Error("Exception from getmailbox powershell", e);
            }
            return "";
        }

        public ADUser GetMtAduser(string customerID, string userPrincipalName)
        {
            var runspace = GetRunspace();
            Pipeline pipeline = runspace.CreatePipeline();
            //pipeline.Commands.AddScript("Get-MTAduser -CustomerID " + customerID + " -UserPrincipalName " + userPrincipalName);
            pipeline.Commands.AddScript("Get-Module");
            ADUser user = new ADUser();
            try
            {
                var results = pipeline.Invoke();
                string errors = GetErrors(pipeline);
                if (errors != "")
                {
                    //maybe fail !
                }

                var stringBuilder = new StringBuilder();
                foreach (PSObject obj in results)
                {
                    stringBuilder.AppendLine(obj.ToString() + "\n");
                    foreach (PSPropertyInfo info in obj.Properties)
                    {
                        stringBuilder.AppendLine(info.ToString() + "\n");
                    }
                }
                log.Info(stringBuilder.ToString());

/*                foreach (PSObject obj in results)
                {
                    user = ADUser.GetAdUser(obj);                    
                    log.Info(user.ToString() + "\n");
                }*/
            }
            catch (Exception e)
            {
                log.Error("Exception from getmailbox powershell", e);
            }
            return user;
        }

        public ADUser[] GetMtAdusers(string customerID)
        {
            var runspace = GetRunspace();
            Pipeline pipeline = runspace.CreatePipeline();
            pipeline.Commands.AddScript("Get-MTAduser -CustomerID " + customerID);
            ADUser[] users = null;
            try
            {
                var results = pipeline.Invoke();
                string errors = GetErrors(pipeline);
                if (errors != "")
                {
                    //maybe fail !
                }
                
                foreach (PSObject obj in results)
                {
                    if (obj.Properties != null && obj.Properties["UserPrincipalName"] != null && obj.Properties["UserPrincipalName"].Value != null)
                    {
                        if (obj.Properties["UserPrincipalName"].Value is string)
                        {
                            users = new ADUser[1];
                            users[0] = ADUser.GetAdUser(obj);
                        }
                        else if (obj.Properties["UserPrincipalName"].Value is string[])
                        {
                            int numberOfUsers = ((Array)obj.Properties["UserPrincipalName"].Value).Length;
                            users = new ADUser[numberOfUsers];
                            for (int i = 0; i < numberOfUsers; i++)
                            {
                                users[i] = ADUser.GetAdUser(obj, i);
                            }
                        }
                    }
                }

                if(users != null)
                {
                    foreach (ADUser user in users)
                    {
                        log.Info(user.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                log.Error("Exception from getmailbox powershell", e);
            }
            return users;
        }

        public string SetMtAduser(ADUser adUser)
        {
            var runspace = GetRunspace();
            Pipeline pipeline = runspace.CreatePipeline();
            pipeline.Commands.AddScript("Set-MTAduser " + adUser.GetPSParameters());

            try
            {
                var results = pipeline.Invoke();
                string errors = GetErrors(pipeline);
                if (errors != "")
                {
                    //maybe fail !
                }

                // convert the script result into a single string
                var stringBuilder = new StringBuilder();
                foreach (PSObject obj in results)
                {
                    ADUser user = ADUser.GetAdUser(obj);
                    stringBuilder.AppendLine(user.ToString() + "\n");

                    /*stringBuilder.AppendLine(obj.ToString() + "\n");
                    foreach (PSPropertyInfo info in obj.Properties)
                    {
                        stringBuilder.AppendLine(info.ToString() + "\n");
                    }
                    string distinguishedName = GetString(obj, "DistinguishedName");
                    stringBuilder.AppendLine("Distinguished Name: " + distinguishedName + "\n");

                    foreach (Object alias in GetHashSet(obj, "AddedProperties"))
                    {
                        stringBuilder.AppendLine("Added Properties: " + alias);
                    }*/
                }

                return stringBuilder.ToString();
            }
            catch (Exception e)
            {
                log.Error("Exception from getmailbox powershell", e);
            }
            return "";
        }

        public static string EnableMailbox(string accountName)
        {
            var runspace = GetRunspace();

            Pipeline pipeline = runspace.CreatePipeline();

            pipeline.Commands.AddScript("Enable-DETMailbox -Name " + accountName + " -Type Shared -Capacity 5");

            var results = pipeline.Invoke();
            var stringBuilder = new StringBuilder();
            foreach (PSObject obj in results)
            {
                stringBuilder.AppendLine(obj.ToString());
            }

            return stringBuilder.ToString();
        }

        //Clean up after ourselves
        public void Dispose()
        {
            CloseRunspace();
        }

        private static string GetErrors(Pipeline pipeline)
        {
            StringBuilder sb = new StringBuilder();
            if (pipeline.Error.Count > 0)
            {
                var error = pipeline.Error.Read() as Collection<ErrorRecord>;
                if (error != null)
                {
                    foreach (ErrorRecord er in error)
                    {
                        log.Error("Exception from PS: ", er.Exception);
                        sb.Append(er.Exception.Message);
                    }
                }
            }
            return sb.ToString();
        }

        private bool GetBoolean(PSObject obj, string name)
        {
            if (!IsHashRefNull(obj, name))
            {
                return (bool)obj.Properties[name].Value;
            }
            return false;
        }

        private string GetString(PSObject obj, string name)
        {
            if (!IsHashRefNull(obj, name))
            {
                return (string)obj.Properties[name].Value;
            }
            return "";
        }

        private Object[] GetObjectArray(PSObject obj, string name)
        {
            if (!IsHashRefNull(obj, name))
            {
                return (Object[])obj.Properties[name].Value;
            }
            return new Object[] {};
        }

        private HashSet<string> GetHashSet(PSObject obj, string name)
        {
            if (!IsHashRefNull(obj, name))
            {
                return (HashSet<string>)obj.Properties[name].Value;
            }
            return new HashSet<string>();
        }

        private SortedDictionary<string, string> GetSortedDictionary(PSObject obj, string name)
        {
            if (!IsHashRefNull(obj, name))
            {
                return (SortedDictionary<string, string>)obj.Properties[name].Value;
            }
            return new SortedDictionary<string, string>();
        }

        private Boolean IsHashRefNull(PSObject obj, string name)
        {
            return obj.Properties[name] == null || obj.Properties[name].Value == null;
        }
    }
}
