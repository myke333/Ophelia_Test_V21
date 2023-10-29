
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Remote;
using Ophelia_Test_V21.AFuncionesGlobales;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using ParamAccessHelper;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21
{
    class AbrirPrograma
    {
        protected const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        //private const string WpfAppId = @"C:\Users\mpagani\Source\Samples\WpfUiTesting\WpfUiTesting.Wpf\bin\Debug\WpfUiTesting.Wpf.exe";
        protected static WindowsDriver<WindowsElement> session;
        protected static WindowsDriver<WindowsElement> ApplicationSession;
        protected static WindowsDriver<WindowsElement> sessionFinal;
        static public object[] Pdatain;
        static public object[] Pdatain2;
        static public object[] pdataprocesar;
        string Code;
        string token;

        string usuer;
        System.Diagnostics.Process p = new System.Diagnostics.Process();
        string UserDomains;
        string PassDomains;

        public AbrirPrograma()
        {
      
        }


        public AbrirPrograma(string CodigoPrograma, string usuario, string UserDomain = null, string PassDomain = null)
        {

            this.Code = CodigoPrograma;
            this.usuer = usuario;
            this.UserDomains = UserDomain;
            this.PassDomains = PassDomain;

        }

        public static string SelectNavigator()
        {
            //navigator Es el Valor que Indica en que navegador abre el proyecto IE (Internet Explorer) - Edge (Microsoft Edge)
            var navigator = "Edge";
            var valueNavigator = "";
            if (navigator == "IE")
            {
                valueNavigator = "IE";
            }
            else if(navigator == "Edge")
            {
                valueNavigator = "Edge";
            }
            return valueNavigator;
        }


        public WindowsDriver<WindowsElement> Start2(string database, string ambiente, string bandera = null)
        {
            var appOptions = new AppiumOptions();
            try
            {

                string encript;
                string host = Dns.GetHostName();
                IPAddress[] iPs = (Dns.GetHostAddresses(host));
                List<string> ip = new List<string>();
                foreach (IPAddress addr in iPs)
                {
                    if (addr.AddressFamily == AddressFamily.InterNetwork)
                    {
                        ip.Add(addr.ToString());
                    }
                }
                //ipf = GetLocalIpAddress();
                Ophelia.Proteccion.Motor a = new Ophelia.Proteccion.Motor();
                Ophelia.Proteccion.API32 vProceso = new Ophelia.Proteccion.API32(database);
                encript = vProceso.FStGn_Encripta(this.usuer + ";" + this.Code, a);
                foreach (string ipf in ip)
                {

                    SqlAdapter.UdpOPSessioV2(this.usuer, ipf, encript, database);
                }
                //KACTUS_UI_TEST.ServiceHost.ServiceGeneralClient Logear = new KACTUS_UI_TEST.ServiceHost.ServiceGeneralClient();
                //KACTUS_UI_TEST.ServiceHost.RequestMessage m = new KACTUS_UI_TEST.ServiceHost.RequestMessage();
                //m.Producto = "Kactus";
                //m.Usuario = this.usuer;
                //m.Contrasena = "abc123";
                //KACTUS_UI_TEST.ServiceHost.ResponseMessage resta = Logear.Loguear(m);

                object[] Pdatain2 = new object[8];
                string Pass;
                if (this.usuer == "CAROR621" || this.usuer == "ODESAR" || this.usuer == "OAUDI")
                {
                    Pass = "abc123";
                }
                else if (this.usuer == "KAUTO")
                {
                    Pass = "Abcd1234*";
                }
                else if (this.usuer == "DRoselyC")
                {
                    Pass = "C181176125131126127183192";
                }
                else if (this.usuer == "GRANGOSE")
                {
                    Pass = "C143145127119129126144142";
                }
                else if (this.usuer == "BATBTA")
                {
                    Pass = "BATBTA";
                }
                else if (this.usuer == "KACTUS")
                {
                    Pass = "TEMPO10";
                }
                else if (this.usuer == "KADMOTUY")
                {
                    Pass = "C143145127129128126144142";
                }
                else if (this.usuer == "TPRUEBAS" || this.usuer == "KOBA" || this.usuer == "KAUTONE" || this.usuer == "UCalida1")
                {
                    Pass = "Medi20sp";
                }
                else if (this.usuer == "KNEAUTOM")
                {
                    Pass = "Medi22sp";
                }
                else if (this.usuer == "CINDYD")
                {
                    Pass = "col2020**";
                }
                else if (this.usuer == "KADMON" || this.usuer == "DWKE" || this.usuer == "DWKEORA")
                {
                    Pass = "C143145127129128126144142";
                }
                else if (this.usuer == "CFBALLEN")
                {
                    Pass = "abc123";
                }
                else if (this.usuer == "KCT_SVPK")
                {
                    Pass = "KCT_SVPK";
                }
                else if (this.usuer == "KDemos")
                {
                    Pass = "ABCD1234";
                }
                else if (this.usuer == "KPREPAP")
                {
                    Pass = "KPREPAP";
                }
                else if (this.usuer == "KPREPA")
                {
                    Pass = "KPREPA";
                }
                else if (this.usuer == "UNatalia")
                {
                    Pass = "Medi20sp";
                }
                else if (this.usuer == "ONATALIA")
                {
                    Pass = "ONATALIA";
                }
                else if (this.usuer == "KFABILU")
                {
                    Pass = "C143145127129128126144142";
                }
                else if (this.usuer == "KSDesar")
                {
                    Pass = "Medi22sp";
                }
                else
                {
                    Pass = "C175126128127176174";
                }
                Pdatain2[0] = this.usuer;
                Pdatain2[1] = Pass;
                Pdatain2[2] = "127.0.0.1";
                Pdatain2[3] = "18832733-B389-4428-8A1C-742FD94943EA";
                Pdatain2[4] = this.Code;
                Pdatain2[5] = this.usuer;
                Pdatain2[6] = 421;
                Pdatain2[7] = "";

                object[] Pdata = new object[1];
                int ResUrl = 0;
                if (usuer == "CFBALLEN" || usuer == "KCT_SVPK" || usuer == "BATBTA")
                {
                    //System.Diagnostics.//Debugger.Launch();
                    Ophelia_Test_V21.SmenWSp105.KGnSmenwWCFClient Url = new Ophelia_Test_V21.SmenWSp105.KGnSmenwWCFClient();
                    ResUrl = Url.ExecRemoteFunction("GenerarLogEntrada", Pdatain2, out object[] Pdata3, out string PstError2);
                    Pdata = Pdata3;
                }
                else if (usuer == "KTEST" || usuer == "KDemos" || usuer == "KSDesar")
                {
                    //Debugger.Launch();
                    Ophelia_Test_V21.SmenwSPEnv.KGnSmenwWCFClient Url = new Ophelia_Test_V21.SmenwSPEnv.KGnSmenwWCFClient();
                    ResUrl = Url.ExecRemoteFunction("GenerarLogEntrada", Pdatain2, out object[] Pdata3, out string PstError2);
                    Pdata = Pdata3;
                }
                else if (usuer == "UNatalia" || usuer == "ONATALIA" || usuer == "KOBA" || usuer == "KAUTONE")
                {
                    Ophelia_Test_V21.SmenwSPEnv2.KGnSmenwWCFClient Url = new Ophelia_Test_V21.SmenwSPEnv2.KGnSmenwWCFClient();
                    ResUrl = Url.ExecRemoteFunction("GenerarLogEntrada", Pdatain2, out object[] Pdata3, out string PstError2);
                    Pdata = Pdata3;
                }
                else if (usuer == "KNEAUTOM" || usuer == "KPREPAP" || usuer == "KPREPA" || usuer == "TPRUEBAS" || usuer == "UCalida1")
                {
                    Ophelia_Test_V21.SmenwSPEnv5.KGnSmenwWCFClient Url = new Ophelia_Test_V21.SmenwSPEnv5.KGnSmenwWCFClient();
                    ResUrl = Url.ExecRemoteFunction("GenerarLogEntrada", Pdatain2, out object[] Pdata3, out string PstError2);
                    Pdata = Pdata3;
                }
                else if (usuer == "KADMON")
                {
                    Ophelia_Test_V21.SmenwSPEnv3.KGnSmenwWCFClient Url = new Ophelia_Test_V21.SmenwSPEnv3.KGnSmenwWCFClient();
                    ResUrl = Url.ExecRemoteFunction("GenerarLogEntrada", Pdatain2, out object[] Pdata3, out string PstError2);
                    Pdata = Pdata3;
                }
                else if (usuer == "KFABILU")
                {
                    Ophelia_Test_V21.SmenwSPEnv4.KGnSmenwWCFClient Url = new Ophelia_Test_V21.SmenwSPEnv4.KGnSmenwWCFClient();
                    ResUrl = Url.ExecRemoteFunction("GenerarLogEntrada", Pdatain2, out object[] Pdata3, out string PstError2);
                    Pdata = Pdata3;
                }
                else if (usuer == "DWKE" || usuer == "DWKEORA")
                {
                    Ophelia_Test_V21.SmenwErrores.KGnSmenwWCFClient Url = new Ophelia_Test_V21.SmenwErrores.KGnSmenwWCFClient();
                    ResUrl = Url.ExecRemoteFunction("GenerarLogEntrada", Pdatain2, out object[] Pdata3, out string PstError2);
                    Pdata = Pdata3;
                }
                else
                {
                    //System.Diagnostics.//Debugger.Launch();
                    Ophelia_Test_V21.KGnSmenw2.KGnSmenwWCFClient Url = new Ophelia_Test_V21.KGnSmenw2.KGnSmenwWCFClient();
                    ResUrl = Url.ExecRemoteFunction("GenerarLogEntrada", Pdatain2, out object[] Pdata3, out string PstError2);
                    Pdata = Pdata3;
                }
                //Debugger.Launch();
                
                if (ResUrl == 0)
                {
                    //Debugger.Launch();
                    string navigator = SelectNavigator();
                    if (navigator == "Edge")
                    {
                        //Modo Edge
                        string UrlProgramas = "" + " " + "\"" + Pdata[0].ToString() + "\"";
                        appOptions.AddAdditionalCapability("platformName", @"Windows");
                        appOptions.AddAdditionalCapability("deviceName", @"WindowsPC");
                        appOptions.AddAdditionalCapability("ms:experimental-webdriver", true);
                        appOptions.AddAdditionalCapability("ms:waitForAppLaunch", "1");
                        appOptions.AddAdditionalCapability("app", @"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe");
                        appOptions.AddAdditionalCapability("appArguments", UrlProgramas);
                        appOptions.AddAdditionalCapability("platformVersion", "10");
                    }
                    else if(navigator == "IE")
                    {
                        //Modo IE
                        string UrlProgramas = Pdata[0].ToString();
                        appOptions.AddAdditionalCapability("platformName", @"Windows");
                        appOptions.AddAdditionalCapability("deviceName", @"WindowsPC");
                        appOptions.AddAdditionalCapability("ms:experimental-webdriver", true);
                        appOptions.AddAdditionalCapability("ms:waitForAppLaunch", "1");
                        appOptions.AddAdditionalCapability("app", @"C:\Program Files (x86)\Internet Explorer\iexplore.exe");
                        appOptions.AddAdditionalCapability("appArguments", UrlProgramas);
                        appOptions.AddAdditionalCapability("platformVersion", "10");
                    }


                    ////Debugger.Launch();
                    session = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appOptions);
                    session.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(15);
                    Thread.Sleep(15000);
                    session.Manage().Window.Maximize();
                    Thread.Sleep(14000);

                    if (this.usuer == "CFBALLEN" || this.usuer == "KCT_SVPK")
                    {
                        AFuncionesGlobales.newFunctionsAppium.newFunctions_1.ScreenshotNuevoSP105(this.Code, database);
                    }
                    Thread.Sleep(2000);
                    
                    if (bandera == null)
                    {
                        if (this.Code == "KNmContr")
                        {
                            WindowsDriver<WindowsElement> rootSession = null;
                            rootSession = PruebaCRUD.RootSession();
                            rootSession = PruebaCRUD.ReloadSession(rootSession, "TFrmXP");
                            rootSession.Keyboard.SendKeys(Keys.Enter);
                        }
                        if (navigator == "IE")
                        {
                            ////Se comenta en modo EDGE
                            var ApplicationWindow = session.FindElementByClassName("IEFrame");
                            var ApplicationSessionHandle = ApplicationWindow.GetAttribute("NativeWindowHandle");
                            ApplicationSessionHandle = (int.Parse(ApplicationSessionHandle)).ToString("x"); //Convert to hex
                            var appCapabilities = new AppiumOptions();
                            appCapabilities.AddAdditionalCapability("appTopLevelWindow", ApplicationSessionHandle);
                            Thread.Sleep(5000);
                            var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
                            if (this.Code == "KNmAumge")
                            {
                                new Actions(ApplicationSession).MoveToElement(ApplicationWindow, 0, 0).MoveByOffset(ApplicationWindow.Size.Width / 2, ApplicationWindow.Size.Height / 2).ContextClick().Perform();
                                ApplicationSession.Keyboard.SendKeys(Keys.Enter);
                            }
                            return ApplicationSession;
                        }
                    }
                    sessionFinal = session;
                    if (navigator == "IE")
                    {
                        sessionFinal = ApplicationSession;
                    }
                    //return ApplicationSession;
                    return sessionFinal;
                }
                else
                {
                    ////Debugger.Launch();
                    Assert.Fail("Error de respuesta SmenW");
                    return null;
                }
            }
            catch (Exception ex)
            {
                ////Debugger.Launch();
                Console.WriteLine(ex);
                return null;
            }
        }
        public WindowsDriver<WindowsElement> ReloadSession(WindowsDriver<WindowsElement> session)
        {
            var ApplicationWindow = session.FindElementByClassName("IEFrame");
            var ApplicationSessionHandle = ApplicationWindow.GetAttribute("NativeWindowHandle");
            ApplicationSessionHandle = (int.Parse(ApplicationSessionHandle)).ToString("x"); //Convert to hex

            var appCapabilities = new AppiumOptions();
            //appCapabilities.AddAdditionalCapability("app", "Root");
            appCapabilities.AddAdditionalCapability("appTopLevelWindow", ApplicationSessionHandle);
            var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
            return ApplicationSession;
        }


        public WindowsDriver<WindowsElement> RootSession()
        {
         
            var appCapabilities = new AppiumOptions();
            appCapabilities.AddAdditionalCapability("app", "Root");
            var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
            return ApplicationSession;
        }

        public void Stop()
        {
            string navigator = SelectNavigator();
            string varNavigator = null;
            if (navigator == "Edge")
            {
                //Modo Edge
                varNavigator = "msedge";
            }
            else if (navigator == "IE")
            {
                //Modo IE
                varNavigator = "iexplore";
            }
            
            Process[] processes = Process.GetProcessesByName(varNavigator);


            if (processes.Length > 0)
            {
                for (int i = 0; i < processes.Length; i++)
                {
                    processes[i].Kill();
                }
            };

            //BrowserWindow.ClearCache();
            //BrowserWindow.ClearCookies();
        }
        public static string GetLocalIpAddress()
        {
            UnicastIPAddressInformation mostSuitableIp = null;

            var networkInterfaces = NetworkInterface.GetAllNetworkInterfaces();

            foreach (var network in networkInterfaces)
            {
                if (network.OperationalStatus != OperationalStatus.Up)
                    continue;

                var properties = network.GetIPProperties();

                if (properties.GatewayAddresses.Count == 0)
                    continue;

                foreach (var address in properties.UnicastAddresses)
                {
                    if (address.Address.AddressFamily != AddressFamily.InterNetwork)
                        continue;

                    if (IPAddress.IsLoopback(address.Address))
                        continue;

                    if (!address.IsDnsEligible)
                    {
                        if (mostSuitableIp == null)
                            mostSuitableIp = address;
                        continue;
                    }

                    // The best IP is the IP got from DHCP server
                    if (address.PrefixOrigin != PrefixOrigin.Dhcp)
                    {
                        if (mostSuitableIp == null || !mostSuitableIp.IsDnsEligible)
                            mostSuitableIp = address;
                        continue;
                    }

                    return address.Address.ToString();
                }
            }

            return mostSuitableIp != null
             ? mostSuitableIp.Address.ToString()
                : "";
        }
    }

    class TFSData
    {
        private string testcaseID;
        private string suites;
        private string plans;
        private string ServerURL = ConfigurationManager.AppSettings["ServerURL"];
        private string ServerPat = ConfigurationManager.AppSettings["ServerPat"];

        public TFSData(string testcaseIDs, string plan = null, string suite = null)
        {
            this.testcaseID = testcaseIDs;
            this.suites = plan;
            this.plans = suite;
        }

        public DataSet GetParams()
        {
            //System.Diagnostics.//Debugger.Launch();
            DataSet ds = new DataSet();
            GetTestCaseParams p = new GetTestCaseParams();
            p.VstsURI = ServerURL;
            p.Pat = ServerPat;

            Task.Run(async () => { ds = await p.GetParams(this.testcaseID); }).GetAwaiter().GetResult();
            return ds;
        }




        public DataSet GetParamsExecutionCases()
        {
            DataSet ds = new DataSet();
            GetTestCaseParams p = new GetTestCaseParams();
            p.VstsURI = ServerURL;
            p.Pat = ServerPat;

            Task.Run(async () => { ds = await p.GetEcecutionTfsTestCase(this.plans, this.suites); }).GetAwaiter().GetResult();
            return ds;
        }


        public DataSet GetParamsBuildTest()
        {
            DataSet ds = new DataSet();
            GetTestCaseParams p = new GetTestCaseParams();
            p.VstsURI = ServerURL;
            p.Pat = ServerPat;

            Task.Run(async () => { ds = await p.GetBuildTest(); }).GetAwaiter().GetResult();
            return ds;
        }


        public List<string> GetQuery(string q)
        {
            List<string> ds = new List<string>();
            GetTestCaseParams p = new GetTestCaseParams();
            p.VstsURI = ServerURL;
            p.Pat = ServerPat;

            Task.Run(async () => { ds = await p.GetTestCasesByQuery(q); }).GetAwaiter().GetResult();
            return ds;
        }


        ~TFSData()
        {

        }

    }
}
