using System;

using System.Collections.Generic;

using System.Collections.ObjectModel;

using System.Linq;

using System.Management.Automation;

using System.Management.Automation.Runspaces;

using System.Text;

using System.Threading.Tasks;



namespace DetePoc.Provisioning.PowerShell
{

    public class LocalExchange
    {

        private static Runspace _runspace = null;

        private static object _runspaceLock = new object();

        private static object _runspaceStateLock = new object();



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



                        initial.ImportPSModule(new string[] { "ActiveDirectory", "DETManagement" });



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

                        _runspace.Open();





                        RunspaceInvoke invoker = new RunspaceInvoke(_runspace);

                        invoker.Invoke("Set-ExecutionPolicy Bypass");



                        Pipeline pipeline = _runspace.CreatePipeline();

                        pipeline.Commands.AddScript("$s = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://eq001yex.CIDPOC.Server-Complex.com/PowerShell/ -Authentication Kerberos", false);

                        pipeline.Commands.AddScript("Import-PSSession $s");



                        pipeline.Invoke();



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



        public static string GetMailbox()
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

    }

}

