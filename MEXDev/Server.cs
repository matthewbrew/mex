using System;
using System.Collections;
using System.IO;
using System.Runtime.Remoting;
using System.Runtime.Remoting.Channels;
using Ch.Elca.Iiop;
using Ch.Elca.Iiop.Services;
using log4net;
using Microsoft.Win32;
using omg.org.CORBA;
using omg.org.CosNaming;
using omg.org.CosTrading;
using omg.org.CosTradingRepos;
using omg.org.CosTradingRepos.ServiceTypeRepository_package;
using _Property=omg.org.CosTrading._Property;
using TypeCode=omg.org.CORBA.TypeCode;

namespace AutomationServer
{
    public class Server
    {
        private static readonly ILog log = LogManager.GetLogger(typeof (Server));
        

        public static RegistryKey generalKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Automation\properties");
        public static string CURRENTSERVICE = (string)generalKey.GetValue("CurrentService");
        public static string IDL;

        public static string MACHINEIP;

        // general properties
        public static string NAMINGSERVICEIP = (string)generalKey.GetValue("NamingServiceIp", "corba.dev.netregistry.net");

        public static int NAMINGSERVICEPORT = int.Parse((string) generalKey.GetValue("NamingServicePort", "2001"));
        public static string OFFERFILEPATH;
        public static int PORT;
        public static RegistryKey serverKey;

        // service specific properties
        public static string SERVICENAME;

        public static string TRADINGNAME;

        public static string TRADINGPROPERTY;

        public static string TRADINGPROPERTYVALUE;
        public static string TRADINGSERVICEIP = (string)generalKey.GetValue("TradingServiceIp", "corba.dev.netregistry.net");
        public static int TRADINGSERVICEPORT = int.Parse((string) generalKey.GetValue("TradingServicePort", "2011"));
        public static string WEBSERVERNAME;

        private readonly MarshalByRefObject automationServer;

        private Lookup lookup;

        //override the CURRENTSERVICE value
        public Server() : this(CURRENTSERVICE)
        {
        }

        public Server(string currentService)
        {
            serverKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Automation\" + currentService);

            SERVICENAME = (string) serverKey.GetValue("ServiceName", "MEX");


            WEBSERVERNAME = (string) serverKey.GetValue("webserverName");
            
            Type type = Type.GetType(WEBSERVERNAME);

            if (type == null)
                throw new Exception("Name does not exist: " + WEBSERVERNAME);

            automationServer = (MarshalByRefObject)Activator.CreateInstance(type);

            TRADINGNAME = (string) serverKey.GetValue("TradingName");

            OFFERFILEPATH = (string) serverKey.GetValue("OfferFile");

            TRADINGPROPERTY = (string) serverKey.GetValue("TradingProperty");

            TRADINGPROPERTYVALUE = (string) serverKey.GetValue("TradingPropertyValue");

            IDL = (string) serverKey.GetValue("Idl");

            MACHINEIP = (string) serverKey.GetValue("MachineIP");

            PORT = int.Parse((string) serverKey.GetValue("Port"));

            log.InfoFormat("CORBA SERVER initialising finished!");
        }

        public void run()
        {
            // register the channel
            IDictionary props = new Hashtable();

            IiopChannel chan = null;
            if ((string) generalKey.GetValue("production") == "true")
            {
                props[IiopServerChannel.PORT_KEY] = PORT;
                props[IiopServerChannel.MACHINE_NAME_KEY] = MACHINEIP;
                chan = new IiopChannel(props);
            }
            else
            {
                chan = new IiopChannel(36435);
            }

            ChannelServices.RegisterChannel(chan, false);

            ////string objectURI = "iisserver";
            RemotingServices.Marshal(automationServer);
            //

            CorbaInit init = CorbaInit.GetInit();


            log.InfoFormat("Ready to get NamingContext, IP{0} and PORT {1}", NAMINGSERVICEIP, NAMINGSERVICEPORT);
            log.InfoFormat("Name SERVICENAME " + SERVICENAME);

            NamingContext nameService = init.GetNameService(NAMINGSERVICEIP, NAMINGSERVICEPORT);


            var name = new[] {new NameComponent("automation", "")};
            NamingContext automationContext = null;
            try
            {
                automationContext = nameService.bind_new_context(name);
            }
            catch (Exception e)
            {
                log.Error("Error occurred while binding new context: " + e.Message);
                automationContext = (NamingContext) nameService.resolve(name);
            }

            var components = new[] {new NameComponent(SERVICENAME, "")};
            automationContext.rebind(components, automationServer);

            ORB orb = OrbServices.GetSingleton();


            log.InfoFormat("Ready to get TradingService IP {0} PORT {1}", NAMINGSERVICEIP, NAMINGSERVICEPORT);
            string corbaLoc =
                String.Format("corbaloc::1.2@{0}:{1}/TradingService", TRADINGSERVICEIP, TRADINGSERVICEPORT);
            log.Info("CORBALOC [" + corbaLoc + "]");

            lookup = (Lookup) RemotingServices.Connect(typeof (Lookup), corbaLoc);

            var propTypes = new PropStruct[1];
            propTypes[0] = new PropStruct();
            propTypes[0].name = TRADINGPROPERTY;
            propTypes[0].value_type = orb.create_string_tc(0);
            propTypes[0].mode = PropertyMode.PROP_MANDATORY_READONLY;

            var repos = (ServiceTypeRepository) lookup.type_repos;

            try
            {
                repos.add_type(TRADINGNAME,
                               IDL,
                               propTypes,
                               new string[0]);
            }
            catch (ServiceTypeExists e)
            {
                log.Error("The service type " + TRADINGNAME + " already exists: " + e.Message);
            }

            // try to withdraw the service again in case the server was shutdown unexpectly
            withdrawService();

            TypeCode stringTC = orb.create_string_tc(0);

            var any = new Any(TRADINGPROPERTYVALUE, stringTC);


            var properties = new _Property[1];
            properties[0] = new _Property();
            properties[0].name = TRADINGPROPERTY;
            properties[0].value = any;

            //log.Debug("Exporting service:" + automationServer.get_service_name());
            string s = lookup.register_if.export(automationServer, TRADINGNAME, properties);

            //write trader offer id to a offer file

            var sw = new StreamWriter(OFFERFILEPATH);
            sw.WriteLine(s);
            sw.Close();

            log.Info(SERVICENAME + " is running.");
        }

        public void withdrawService()
        {
            if (File.Exists(OFFERFILEPATH))
            {
                // withdraw offer;

                var sr = new StreamReader(OFFERFILEPATH);
                string id = sr.ReadLine();

                log.Debug("Withdrawing offer: " + id);
                try
                {
                    lookup.register_if.withdraw(id);
                }
                catch (Exception e)
                {
                    log.Error("withdrawing trader offer id '" + id + "' failed: " + e.Message);
                    Console.WriteLine(e.StackTrace);
                }
                sr.Close();

                File.Delete(OFFERFILEPATH);
            }
            //clean up after the automation server
            if (automationServer is IDisposable)
            {
                ((IDisposable)automationServer).Dispose();
            }
        }


        public static void Main(string[] args)
        {
            Console.WriteLine("server started");

            Server server = null;
            if (args.Length == 1)
                server = new Server(args[0]);
            else
                server = new Server();

            server.run();
            Console.ReadLine();
            server.withdrawService();
                        

        }
    }
}