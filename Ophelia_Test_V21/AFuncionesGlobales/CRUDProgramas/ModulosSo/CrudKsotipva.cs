using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSo
{
    class CrudKsotipva : FuncionesVitales
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


        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string file)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (tipo == 1)
            {
                ElementList[3].SendKeys(data[0]);
                ElementList[2].SendKeys(data[1]);
            }
            else
            {
                ElementList[2].Clear();
                ElementList[2].SendKeys(data[2]);
            }
        }

        public static void CrudDetalle1(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string file)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            if (tipo == 1)
            {
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 102, 27);
                desktopSession.Mouse.Click(null);
                for (int i = 0; i < 4; i++)
                {
                    desktopSession.Keyboard.SendKeys(data[i]);
                    if (i <= 2)
                    {
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    }

                }
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
            else
            {
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(data[4]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
        }

        public static void CrudDetalle2(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string file)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            if (tipo == 1)
            {
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 50, 29);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(data[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
            else
            {
                desktopSession.Keyboard.SendKeys(data[1]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
        }

        
        public static void ClickButtonExterno(WindowsDriver<WindowsElement> desktopSession, int tipo)
        {
            //1 -> Dosis 2 -> Cargos
            List<string> errorMessages = new List<string>();
            if (tipo == 1)
            {
                desktopSession.FindElementByName("Dosis").Click();
            }
            else
            {
                desktopSession.FindElementByName("Cargos").Click();
            }
        }
    }
}
