using System;
using System.Collections;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using log4net;
using System.Collections.ObjectModel;
using System.Text;
using System.Management.Automation;
using System.Management.Automation.Host;
using System.Management.Automation.Runspaces;
using System.Security;


namespace Exchange
{
    class Program
    {
        private static Uri uri = new Uri("https://mail.uc-lab.server-secure.com/powershell");
        private static string schemaURI = "http://schemas.microsoft.com/powershell/Microsoft.Exchange";
        private static string user = "GFL\\vlad";
        private static string password = "Password1";
        private static RunspacePool runspace = null;

        private static readonly ILog log = LogManager.GetLogger(typeof(Program));

        static Program()
        {
            runspace = RunspaceFactory.CreateRunspacePool(1, 30, GetConnectionInfo());
        }

        static void Main(string[] args)
        {
            log.Info("Here 111 !!");
            Program prog = new Program();
            using (LocalExchange localExchange = new LocalExchange())
            {
                //log.Info(localExchange.GetMailbox());
                //log.Info(localExchange.GetMailbox());
                //log.Info(localExchange.GetMailbox());
                
                //How to to get help documentation for powershell functions
                //Import-Module testmodule
                //get-help get-mtaduser -full
                log.Info(localExchange.GetMtMailbox());
                log.Info(localExchange.GetMtAduser("vs000013", "mb000013a"));
                log.Info(localExchange.GetMtAdusers("vs000013"));
            }
            Console.Read();
            //LocalExchange.CloseRunspace();
            //prog.Test1();
            //prog.Test1();
            //prog.Test1();
            //prog.Test1();
            //prog.Test1();
            //prog.Test2();
        }

        public void Test2()
        {
            
        }

        public void Test1()
        {
            Collection<PSObject> users = GetUserInformation(10);
            StringBuilder stringBuilder = new StringBuilder();
            log.Info("Here 222 !!");
            if (users != null)
            {
                foreach (PSObject pso in users)
                {
                    stringBuilder.AppendLine(pso.ToString());
                }
            }
            else
            {
                log.Info("Null results");
            }
            log.Info(stringBuilder.ToString());
        }
        
        public Collection<PSObject> GetUserInformation(int count)
        {
            using (PowerShell powershell = PowerShell.Create())
            {
                powershell.AddCommand("Get-Mailbox");
                return invokePowershell(powershell);
            }
        }

        public Collection<PSObject> invokePowershell(PowerShell powershell)
        {
            runspace.Open();
            powershell.RunspacePool = runspace;
            return powershell.Invoke();
        }

        public Runspace GetRunspace()
        {
            //InitialSessionState initialSessionState = InitialSessionState.CreateDefault();
            Runspace runspace = RunspaceFactory.CreateRunspace(GetConnectionInfo());
            //pool.InitialSessionState.ImportPSModule();
            return runspace;
        }

        public static WSManConnectionInfo GetConnectionInfo()
        {            
            SecureString secureStringPassword = new SecureString();
            foreach (char ch in password)
                secureStringPassword.AppendChar(ch);
            PSCredential credential = new PSCredential(user, secureStringPassword); // the password must be of type SecureString
            WSManConnectionInfo connectionInfo = new WSManConnectionInfo(uri, schemaURI, credential);
            connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Basic;
            return connectionInfo;
        }
        
        /*static void Main(string[] args)
        {
            try
            {
                using (RunspacePool rsp = RunspaceFactory.CreateRunspacePool())
                {

                    *//*
                    Runspace myRunspace = RunspaceFactory.CreateRunspace();
                    myRunspace.Open();

                    RunspaceConfiguration rsConfig = RunspaceConfiguration.Create();
                    PSSnapInException snapInException = null;
                    PSSnapInInfo info = rsConfig.AddPSSnapIn("Microsoft.Exchange.Management.PowerShell.Admin", out snapInException);
                    Runspace myRunSpace = RunspaceFactory.CreateRunspace(rsConfig);
                    myRunSpace.Open();
                    Pipeline pipeLine = myRunSpace.CreatePipeline();
                    Command myCommand = new Command("Get-Command");
                    pipeLine.Commands.Add(myCommand);
                    Collection<PSObject> commandResults = pipeLine.Invoke();
                    foreach (PSObject cmdlet in commandResults)
                    {
                        string cmdletName = cmdlet.Properties["Name"].Value.ToString();
                        log.Info(cmdletName);
                    }*//*
                    
                    log.Info("Here !!");
                    log.Info("Here !!");
                    log.Info("Here !!");
                    String ouGroup = "";
                    String ouMBName = "";
                    String domainController = "";
                    
                    rsp.Open();
                    PowerShell ps = PowerShell.Create();
                    ps.RunspacePool = rsp;
                    //get-mailbox -Identity "$()$()\" -ErrorAction silentlycontinue -domaincontroller 
                    //ps.AddScript("get-mailbox -Identity \"$($OUGroup)$($MBName)\" -ErrorAction silentlycontinue -domaincontroller $DomainController");
                    //ps.AddParameter("OUGroup", ouGroup);
                    //ps.AddParameter("MBName", ouMBName);
                    //ps.AddParameter("DomainController", domainController);

                    ps.AddCommand("Get-Command");
                    log.Info(ps.ToString());
                    log.Info("Here !!");
                    StringBuilder stringBuilder = new StringBuilder();
                    Collection<PSObject> results = ps.Invoke();
                    log.Info("Here !!");
                    if (results != null)
                    {
                        foreach (PSObject pso in results)
                        {
                            stringBuilder.AppendLine(pso.ToString());
                        }
                    }
                    else
                    {
                        log.Info("Null results");
                    }
                    log.Info(stringBuilder.ToString());
                }
            }
            catch(Exception e)
            {
                log.Info(e.ToString());
            }
        }*/
    }
}
