using System;
using System.Collections;
using System.IO;
using System.Management.Automation;
using System.Runtime.Serialization.Formatters.Binary;
using automation.exchange;
using Ch.Elca.Iiop.Idl;
using log4net;  
using omg.org.CORBA;
using TypeCode=omg.org.CORBA.TypeCode;
using Exchange;

namespace AutomationServer
{
    [SupportedInterface(typeof(ExchangeHelper))]
    public class Exchange : MarshalByRefObject, ExchangeHelper, IDisposable
    {
        private static readonly ILog log = LogManager.GetLogger(typeof (Exchange));
        private static LocalExchange localExchange = new LocalExchange();


        public Exchange()
        {
        }

        public string get_service_name()
        {
            log.Debug("get_service_name()");
            log.Debug("MEX");
            return "MEX";
        }

        public void createAccount(String name)
        {
            localExchange.GetMailbox();
            /*          String ouGroup = "";
                        String ouMBName = "";
                        String domainController = "";

                        PowerShell ps = PowerShell.Create();
                        ps.AddScript("get-mailbox -Identity \"$($OUGroup)$($MBName)\" -ErrorAction silentlycontinue -domaincontroller $DomainController");
                        ps.AddParameter("OUGroup", ouGroup);
                        ps.AddParameter("MBName", ouMBName);
                        ps.AddParameter("DomainController", domainController);*/

            /*if((get-mailbox -Identity "$($objUser.OUGroup)$($objUser.MBName)" -ErrorAction silentlycontinue -domaincontroller $DomainController) -ne $null)
                { 
                $script:info += "manageUser: ERROR! CREATE command issued but user exists! WTF?!"
                write-host $info[-1] -ForegroundColor black -BackgroundColor Red
                hardexit ("manageUser", $strArgs, $objUser, $script:info, $error, $startTime)
                }

                $script:info += "manageUser: checking if customer exists"
                write-host -ForegroundColor DarkGray $info[-1]
                #if((c:\exchprov\scripts\manageCustomer.ps1 "VERIFY" $objUser.VSName $objUser.DomainSuffix) -eq $false) { 
                if((&"$($homePath)\manageCustomer.ps1" "VERIFY" $objUser.VSName $objUser.DomainSuffix) -eq $false) { 
                    $script:info += "manageUser: Customer doesn't exist! Creating customer.."
                write-host "sleeping first"
                start-sleep -s 5
                    write-host -foregroundcolor yellow $info[-1]
                    #c:\exchprov\scripts\manageCustomer.ps1 "CREATE" $objUser.VSName $objUser.DomainSuffix  
                   &"$($homePath)\manageCustomer.ps1" "CREATE" $objUser.VSName $objUser.DomainSuffix
                }

                # Create user, and set all properties which are not controllable by the customer.
                $script:info += "manageUser: Creating user"
                write-host "sleeping first"
                start-sleep -s 5
                write-host -ForegroundColor DarkGray $info[-1]
                $script:info += new-mailbox -name $objUser.MBName -OrganizationalUnit "$($objUser.OUGroup)" -password (convertto-securestring -string "NoL0g1n4uyet" -AsPlainText -force) `
                        -alias $objUser.MBName -UserPrincipalName "$($objUser.EmailPrefix)@$($objUser.DomainSuffix)" -SamAccountName $objUser.MBName -ResetPasswordOnNextLogon $false -domaincontroller $DomainController
                $script:info += "manageUser: Mailbox created"
                # Set password to not expire
                set-aduser $objUser.MBName -passwordneverexpires $true -server $DomainController
                write-host -ForegroundColor DarkGray $info[-1]*/
            //do stuff
            /*ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            service.Credentials = new WebCredentials("user1@contoso.com", "password");*/
        }

        // This method disables servant unpublishing after timeout (be default timeout is 5 minutes)
        public override object InitializeLifetimeService()
        {
            // live forever
            return null;
        }

        //Clean up after ourselves
        public void Dispose()
        {
            localExchange.Dispose();
        }
    }
}