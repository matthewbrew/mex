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
                            //string importModule = "Import-Module C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\Modules\\MEXModule\\MEXModule.psm1";
                            string importModule = "Import-Module C:\\provisioning\\MEXModule.psm1";
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

        public string GetMtMailbox(string userPrincipalName)
        {
            log.Info("------------------------------------- GET MAILBOX ----------------------------------------");
            var runspace = GetRunspace();
            Pipeline pipeline = runspace.CreatePipeline();
            string command = "Get-MTMailbox -UserPrincipalName " + userPrincipalName;
            pipeline.Commands.AddScript(command);
            PSObjectUtils utils = new PSObjectUtils();
            // convert the script result into a single string
            var stringBuilder = new StringBuilder();
            try
            {
                var results = pipeline.Invoke();
                string errors = GetErrors(pipeline);
                
                foreach (PSObject obj in results)
                {
                    if (utils.IsSuccess(obj))
                    {
                        List<PSObject> data = GetData(obj);
                        foreach (PSObject mailboxObj in data)
                        {
                            stringBuilder.AppendLine(mailboxObj.ToString());
                        }
                        LogResult(utils, obj);
                    }
                    else
                    {
                        LogResult(utils, obj);
                    }
                }
            }
            catch(Exception e)
            {
                log.Error("Exception from getmailbox powershell", e);
                log.Error("Command: " + command);
            }
            return stringBuilder.ToString();
        }

        public ADUser GetMtAduser(string customerID, string userPrincipalName)
        {
            log.Info("------------------------------------- GET ADUSER ----------------------------------------");
            var runspace = GetRunspace();
            Pipeline pipeline = runspace.CreatePipeline();
            string command = "Get-MTAduser -CustomerID " + customerID + " -UserPrincipalName " + userPrincipalName;
            pipeline.Commands.AddScript(command);
            ADUser user = new ADUser();
            PSObjectUtils utils = new PSObjectUtils();
            List<ADUser> users = new List<ADUser>();
            try
            {
                var results = pipeline.Invoke();
                //log any errors from the powershell pipe
                string errors = GetErrors(pipeline);
                foreach (PSObject obj in results)
                {
                    if (utils.IsSuccess(obj))
                    {
                        List<PSObject> data = GetData(obj);
                        foreach (PSObject adUserObj in data)
                        {
                            users.Add(ADUser.GetAdUser(adUserObj));
                        }
                    }
                    else
                    {
                        LogResult(utils, obj);
                    }
                }
            }
            catch (Exception e)
            {
                log.Error("Exception from get-mtaduser powershell", e);
                log.Error("Command: " + command);
            }

            if(users.Count == 1)
            {
                return users[0];
            }
            else if (users.Count > 1)
            {
                log.Error("More than one user found for: " + command);
            }
            else
            {
                log.Error("No user found for: " + command);
            }
            return null;
        }

        public List<ADUser> GetMtAdusers(string customerID)
        {
            log.Info("------------------------------------- GET ADUSERS ----------------------------------------");
            var runspace = GetRunspace();
            Pipeline pipeline = runspace.CreatePipeline();
            string command = "Get-MTAduser -CustomerID " + customerID;
            pipeline.Commands.AddScript(command);
            PSObjectUtils utils = new PSObjectUtils();
            List<ADUser> users = new List<ADUser>();
            try
            {
                var results = pipeline.Invoke();
                //log any errors from the powershell pipe
                GetErrors(pipeline);
                foreach (PSObject obj in results)
                {
                    if (utils.IsSuccess(obj))
                    {
                        List<PSObject> data = GetData(obj);
                        foreach (PSObject adUserObj in data)
                        {
                            users.Add(ADUser.GetAdUser(adUserObj));
                        }
                    }
                    else
                    {
                        LogResult(utils, obj);
                    }                    
                }
            }
            catch (Exception e)
            {
                log.Error("Exception from get-mtaduser powershell", e);
                log.Error("Command: " + command);
            }
            return users;
        }

        public string SetMtAduser(ADUser adUser)
        {
            log.Info("------------------------------------- SET ADUSER ----------------------------------------");
            var runspace = GetRunspace();
            Pipeline pipeline = runspace.CreatePipeline();
            pipeline.Commands.AddScript("New-MTAduser " + adUser.GetPSParameters());
            PSObjectUtils utils = new PSObjectUtils();
            var stringBuilder = new StringBuilder();
            try
            {
                var results = pipeline.Invoke();
                GetErrors(pipeline);
                // convert the script result into a single string
                foreach (PSObject obj in results)
                {
                    if (utils.IsSuccess(obj))
                    {
                        List<PSObject> data = GetData(obj);
                        foreach (PSObject resultObj in data)
                        {
                            stringBuilder.AppendLine(resultObj.ToString());
                        }
                    }
                    else
                    {
                        LogResult(utils, obj);
                    }
                }
            }
            catch (Exception e)
            {
                log.Error("Exception from getmailbox powershell", e);
            }
            return stringBuilder.ToString();
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

        private void LogResult(PSObjectUtils utils, PSObject obj)
        {
            log.Info("Result: " + utils.GetString(obj, "Result") + " ResultMessage: " + utils.GetString(obj, "ResultMessage") + " Error: " + utils.GetString(obj, "ErrorMessage"));
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

        private List<PSObject> GetData(PSObject obj)
        {
            List<PSObject> resultData = new List<PSObject>();
            if (obj.Properties["Data"] == null)
            {
                log.Error("No data returned" + obj);
            }
            else
            {
                Object data = obj.Properties["Data"].Value;
                if (data is PSObject)
                {
                    resultData.Add((PSObject)data);
                }
                else if (data is Collection<PSObject>)
                {
                    foreach (PSObject innerObj in (Collection<PSObject>)data)
                    {
                        resultData.Add(innerObj);
                    }
                }
                else if (data is Object[])
                {
                    Object[] dataArray = (Object[])data;
                    if (dataArray.Length > 0)
                    {
                        if (dataArray[0] is PSObject)
                        {
                            foreach (PSObject innerObj in dataArray)
                            {
                                resultData.Add(innerObj);
                            }
                        }
                        else
                        {
                            log.Error("Array 0 = " + dataArray[0]);
                        }
                    }
                }
                else if (data is Hashtable)
                {
                    Hashtable dataHashtable = (Hashtable)data;
                    if (dataHashtable.Values != null)
                    {
                        PSObject hashConvert = new PSObject();
                        bool didConvert = false;
                        foreach (string key in dataHashtable.Keys)
                        {
                            Object value = dataHashtable[key];
                            if (value is PSObject)
                            {
                                resultData.Add((PSObject)value);
                            }
                            else if (value is Hashtable)
                            {
                                Hashtable innerHashtable = (Hashtable)value;
                                PSObject innerPSObject = new PSObject();
                                foreach (string innerkey in innerHashtable.Keys)
                                {
                                    innerPSObject.Properties.Add(new PSNoteProperty(innerkey, innerHashtable[innerkey]));
                                }
                                resultData.Add(innerPSObject);
                            }
                            else
                            {
                                hashConvert.Properties.Add(new PSNoteProperty(key, dataHashtable[key]));
                                didConvert = true;
                            }
                        }
                        if (didConvert)
                        {
                            resultData.Add(hashConvert);
                        }
                    }
                    else
                    {
                        log.Info("Hashtable empty !! " + dataHashtable);
                    }
                }
                else
                {
                    log.Info("Weird !! " + data);
                }
            }
            return resultData;
        }
    }
}
