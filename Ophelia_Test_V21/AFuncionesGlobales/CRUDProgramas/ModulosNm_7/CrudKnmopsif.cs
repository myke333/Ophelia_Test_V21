using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Drawing;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Ophelia_Test_V21.AFuncionesGlobales;
using OpenQA.Selenium;
using System.Drawing;
using System.Windows.Forms;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System.IO;
using System.Drawing.Imaging;
namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_7
{
    class CrudKnmopsif : FuncionesVitales
    {
        static public WindowsDriver<WindowsElement> RootSession()
        {

            var appCapabilities = new AppiumOptions();
            appCapabilities.AddAdditionalCapability("app", "Root");
            var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
            return ApplicationSession;
        }
        static public WindowsDriver<WindowsElement> ReloadSession(WindowsDriver<WindowsElement> session, string parametro)
        {
            try
            {

                var ApplicationWindow = session.FindElementByClassName(parametro);
                var ApplicationSessionHandle = ApplicationWindow.GetAttribute("NativeWindowHandle");
                ApplicationSessionHandle = (int.Parse(ApplicationSessionHandle)).ToString("x"); //Convert to hex


                var appCapabilities = new AppiumOptions();
                //appCapabilities.AddAdditionalCapability("app", "Root");
                appCapabilities.AddAdditionalCapability("appTopLevelWindow", ApplicationSessionHandle);
                var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);

                string ses = ApplicationSession.PageSource;
                return ApplicationSession;
            }
            catch (Exception ex) { return null; }
        }
        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, string tipo, List<String> data, string file)
        {
            var Elementlist = desktopSession.FindElementsByClassName("TDBEdit");

            var Elementlist1 = desktopSession.FindElementsByName(data[4]);
            //var Elementlist2 = desktopSession.FindElementsByName(data[14]);
            switch (tipo)
            {
                case ("0"):
                    Elementlist[23].SendKeys(data[0]);
                    Elementlist[23].SendKeys(OpenQA.Selenium.Keys.Tab);
                    break;
                case ("1"):
                    Elementlist[21].SendKeys(data[1]);
                    Elementlist[21].SendKeys(OpenQA.Selenium.Keys.Tab);
                    break;
                case ("2"):
                    Elementlist[20].SendKeys(data[2]);
                    break;
                case ("3"):
                    Elementlist[17].SendKeys(data[3]);
                    break;
                case ("4"):
                    Elementlist1[0].Click();
                    break;
                case ("5"):
                    Elementlist[5].SendKeys(data[5]);
                    Elementlist[5].SendKeys(OpenQA.Selenium.Keys.Tab);
                    break;
                case ("6"):
                    Elementlist[4].SendKeys(data[6]);
                    Elementlist[4].SendKeys(OpenQA.Selenium.Keys.Tab);
                    break;
                case ("7"):
                    Elementlist[7].SendKeys(data[7]);
                    Elementlist[7].SendKeys(OpenQA.Selenium.Keys.Tab);
                    break;

                case ("8"):
                    Elementlist[6].SendKeys(data[8]);
                    Elementlist[6].SendKeys(OpenQA.Selenium.Keys.Tab);
                    break;
                case ("9"):
                    Elementlist[19].SendKeys(data[9]);
                    Elementlist[19].SendKeys(OpenQA.Selenium.Keys.Tab);
                    break;
                case ("10"):
                    Elementlist[18].SendKeys(data[9]);
                    Elementlist[18].SendKeys(OpenQA.Selenium.Keys.Tab);
                    break;
                case ("11"):
                    Elementlist[15].SendKeys(data[10]);
                    //Elementlist[15].SendKeys(OpenQA.Selenium.Keys.Tab);
                    break;
                case ("12"):
                    Elementlist[16].SendKeys(data[11]);
                    Elementlist[16].SendKeys(OpenQA.Selenium.Keys.Tab);
                    break;
                case ("13"):
                    Elementlist[3].SendKeys(data[12]);
                    Elementlist[3].SendKeys(OpenQA.Selenium.Keys.Tab);
                    break;
                case ("14"):
                    Elementlist[21].SendKeys(data[14]);
                    Elementlist[21].SendKeys(OpenQA.Selenium.Keys.Tab);
                    break;
            }
            
        }

    }
}
