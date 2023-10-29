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
using System.IO;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSo
{
    class CrudKsotcomi : FuncionesVitales
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
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 148, 32);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(data[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 308, 29);
                desktopSession.Mouse.Click(null);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1000);
                //Nueva Ventana
                WindowsDriver<WindowsElement> rootSession = null;
                rootSession = RootSession();
                rootSession.Keyboard.SendKeys(data[1]);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
            else
            {
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 308, 29);
                desktopSession.Mouse.Click(null);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1000);
                //Nueva Ventana
                WindowsDriver<WindowsElement> rootSession = null;
                rootSession = RootSession();
                for (int i = 0; i < 9; i++)
                {
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                }
                for (int i = 0; i < 9; i++)
                {
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                }
                rootSession.Keyboard.SendKeys(data[2]);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
        }

        public static void CrudDetalle2(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string file)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            if (tipo == 1)
            {
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 39, 25);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(data[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(data[1]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
            else
            {
                desktopSession.Keyboard.SendKeys(data[2]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
        }

        public static void AcomodarDetalle(WindowsDriver<WindowsElement> desktopSession)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            Thread.Sleep(2000);
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 314, 103);
            desktopSession.Mouse.Click(null);
            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowLeft);
            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowLeft);
        }


        public static void ClickButtonExterno(WindowsDriver<WindowsElement> desktopSession, int tipo)
        {
            //1 -> Tipos Representantes 2 -> Población para comité y votaciones 
            List<string> errorMessages = new List<string>();
            if (tipo == 1)
            {
                desktopSession.FindElementByName("Tipos Representantes").Click();
            }
            else
            {
                desktopSession.FindElementByName("Población para comité y votaciones").Click();
            }
        }



    }
}
