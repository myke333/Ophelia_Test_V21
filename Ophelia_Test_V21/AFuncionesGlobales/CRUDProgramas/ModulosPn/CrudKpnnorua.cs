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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosPn
{
    class CrudKpnnorua : FuncionesVitales
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

        public static void CRUD_CONT(WindowsDriver<WindowsElement> desktopSession, string tipo, List<String> data, string file)
        {
            var Elementlist = desktopSession.FindElementsByClassName("TDBEdit");
            if (tipo == "0")
            {
                Elementlist[7].SendKeys(data[0]);
                Elementlist[7].SendKeys(OpenQA.Selenium.Keys.Tab);

            }
        }
        public static void CRUD_ID(WindowsDriver<WindowsElement> desktopSession, string tipo, List<String> data, string file)
        {
            var Elementlist2 = desktopSession.FindElementsByClassName("TDBEdit");
            if (tipo == "1")
            {
                Elementlist2[10].SendKeys(data[1]);
                Elementlist2[10].SendKeys(OpenQA.Selenium.Keys.Tab);
            }
        }
        public static void CRUD_F1(WindowsDriver<WindowsElement> desktopSession, string tipo, List<String> data, string file)
        {
            var Elementlist3 = desktopSession.FindElementsByClassName("TDBEdit");
            if (tipo == "2")
            {
                Elementlist3[6].SendKeys(data[2]);
            }
        }
        public static void CRUD_CN(WindowsDriver<WindowsElement> desktopSession, string tipo, List<String> data, string file)
        {
            var Elementlist4 = desktopSession.FindElementsByClassName("TComboBox");
            if (tipo == "3")
            {
                Elementlist4[4].SendKeys(data[3]);
                Elementlist4[4].SendKeys(OpenQA.Selenium.Keys.Tab);
            }
            else
            {
                if (tipo == "5")
                {
                    Elementlist4[4].SendKeys(data[5]);
                    Elementlist4[4].SendKeys(OpenQA.Selenium.Keys.Tab);
                }
            }
        }
        public static void CRUD_F2(WindowsDriver<WindowsElement> desktopSession, string tipo, List<String> data, string file)
        {
            var Elementlist5 = desktopSession.FindElementsByClassName("TDBEdit");
            if (tipo == "4")
            {
                Elementlist5[28].SendKeys(data[4]);
            }
        }
    }
}
