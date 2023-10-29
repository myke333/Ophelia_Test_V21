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
    class CrudKnmmctic : FuncionesVitales
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
            switch (tipo)
            {
                case ("0"):
                    Elementlist[3].SendKeys(data[0]);
                    Elementlist[3].SendKeys(OpenQA.Selenium.Keys.Tab);
                    Elementlist[2].SendKeys(data[1]);
                    break;
                case ("1"):
                    Elementlist[2].Clear();
                    Elementlist[2].SendKeys(data[2]);
                    break;
            }
        }
        public static void CRUD_DETALLE(WindowsDriver<WindowsElement> desktopSession, string tipo, List<String> data, string file)
        {
            var Elementlist = desktopSession.FindElementsByClassName("TKactusDBgrid");
            switch (tipo)
            {
                case ("0"):
                    desktopSession.Mouse.MouseMove(Elementlist[1].Coordinates, 348, 23);
                    desktopSession.Mouse.Click(null);
                    desktopSession.Keyboard.SendKeys(data[0]);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    break;
                case ("1"):
                    desktopSession.Mouse.MouseMove(Elementlist[1].Coordinates, 348, 23);
                    desktopSession.Mouse.Click(null);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Delete);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Delete);
                    desktopSession.Keyboard.SendKeys(data[1]);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    break;
            }
        }
    }
}