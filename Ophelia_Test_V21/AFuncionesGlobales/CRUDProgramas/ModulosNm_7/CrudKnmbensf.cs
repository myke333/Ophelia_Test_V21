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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_7
{
    class CrudKnmbensf : FuncionesVitales
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
            var Elementlist1 = desktopSession.FindElementsByName(data[3]);
            var Elementlist2 = desktopSession.FindElementsByName(data[7]);
            var Elementlist3 = desktopSession.FindElementsByName(data[8]);
            var Elementlist4 = desktopSession.FindElementsByName(data[11]);
            switch (tipo)
            {
                case ("0"):
                    Elementlist[15].SendKeys(data[0]);
                    Elementlist[14].SendKeys(data[1]);
                    Elementlist[14].SendKeys(OpenQA.Selenium.Keys.Tab);
                    Elementlist[11].SendKeys(data[2]);
                    break;
                case ("1"):
                    Elementlist1[0].Click();
                    Elementlist[9].SendKeys(data[4]);
                    Elementlist[9].SendKeys(OpenQA.Selenium.Keys.Tab);
                    Elementlist[6].SendKeys(data[5]);
                    break;
                case ("2"):
                    Elementlist[5].SendKeys(data[6]);
                    Elementlist[5].SendKeys(OpenQA.Selenium.Keys.Tab);
                    Elementlist2[0].Click();
                    break;
                case ("3"):
                    Elementlist3[0].Click();
                    Elementlist[3].SendKeys(data[9]);
                    Elementlist[2].SendKeys(data[10]);
                    break;
                case ("4"):
                    Elementlist4[0].Click();
                    break;
            }
        }
    }
}