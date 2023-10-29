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
    class CrudKnmmcmac : FuncionesVitales
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
            var Elementlist2 = desktopSession.FindElementsByClassName("TTabSheet");
            switch (tipo)
            {
                case ("0"):
                    Elementlist[18].SendKeys(data[0]);
                    Elementlist[19].SendKeys(data[1]);
                    Elementlist[19].SendKeys(OpenQA.Selenium.Keys.Tab);
                    Elementlist[12].SendKeys(data[2]);
                    Elementlist[12].SendKeys(OpenQA.Selenium.Keys.Tab);
                    Elementlist[10].SendKeys(data[3]);
                    Elementlist[10].SendKeys(OpenQA.Selenium.Keys.Tab);
                    Elementlist[8].SendKeys(data[3]);
                    Elementlist[8].SendKeys(OpenQA.Selenium.Keys.Tab);
                    Elementlist[3].SendKeys(data[4]);
                    Elementlist[3].SendKeys(OpenQA.Selenium.Keys.Tab);
                    Elementlist[5].SendKeys(data[5]);
                    desktopSession.FindElementByName("Novedad").Click();
                    desktopSession.Mouse.MouseMove(Elementlist2[0].Coordinates, 46, 60);
                    desktopSession.Mouse.Click(null);
                    desktopSession.Keyboard.SendKeys(data[6]);
                    desktopSession.Mouse.MouseMove(Elementlist2[0].Coordinates, 163, 60);
                    desktopSession.Mouse.Click(null);
                    desktopSession.Keyboard.SendKeys(data[4]);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    break;
                case ("1"):
                    Elementlist[5].Clear();
                    Elementlist[5].SendKeys(data[7]);
                    break;
            }
        }
        public static void CRUD_DETALLE(WindowsDriver<WindowsElement> desktopSession, string tipo, List<String> data, string file)
        {
            var Elementlist = desktopSession.FindElementsByClassName("TKactusDBgrid");
            switch (tipo)
            {
                case ("0"):
                    desktopSession.FindElementByName("Anexos de comisiones").Click();
                    break;
                case ("1"):
                    desktopSession.Mouse.MouseMove(Elementlist[1].Coordinates, 153, 30);
                    desktopSession.Mouse.Click(null);
                    desktopSession.Keyboard.SendKeys(data[0]);
                    desktopSession.Mouse.MouseMove(Elementlist[1].Coordinates, 315, 30);
                    desktopSession.Mouse.Click(null);
                    desktopSession.Keyboard.SendKeys(data[1]);
                    break;
                case ("2"):
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Delete);
                    desktopSession.Mouse.MouseMove(Elementlist[1].Coordinates, 315, 30);
                    desktopSession.Mouse.Click(null);
                    desktopSession.Keyboard.SendKeys(data[2]);
                    break;
            }
        }
    }
}
