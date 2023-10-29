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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosGn
{
    class CrudKgnmseri : FuncionesVitales
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
                ElementList[4].SendKeys(data[0]);
                ElementList[3].SendKeys(data[1]);
                ElementList[2].SendKeys(data[2]);
            }
            else
            {
                ElementList[3].Clear();
                ElementList[3].SendKeys(data[3]);
            }
        }

        public static void CrudDetalle1(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string file)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TKactusDBgrid");
           
            if (tipo == 1)
            {
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, 113, 35);
                desktopSession.Mouse.Click(null);
                for (int i = 0; i < 2; i++)
                {
                    desktopSession.Keyboard.SendKeys(data[i]);
                    Thread.Sleep(500);
                    if (i == 1) { desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter); break; }
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                }
            } 
            else
            {
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, 82, 44);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(data[2]);;
            }
        }

        public static void CrudDetalle2(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string file)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TKactusDBgrid");
           
            if (tipo == 1)
            {
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, 110, 26);
                desktopSession.Mouse.Click(null);
                for (int i = 0; i < 3; i++)
                {
                    desktopSession.Keyboard.SendKeys(data[i]);
                    Thread.Sleep(500);
                    if (i == 2) { desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter); break; }
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                }
            }
            else
            {
                Thread.Sleep(2000);
                desktopSession.Keyboard.SendKeys(data[3]);
            }
        }

      public static void ClickButton(WindowsDriver<WindowsElement> desktopSession, int tipo)
{
    List<string> errorMessages = new List<string>();
    WebDriverWait wait = new WebDriverWait(desktopSession, TimeSpan.FromSeconds(10));

    if (tipo == 1)
    {
        IWebElement buttonElement = wait.Until(ExpectedConditions.ElementToBeClickable(By.Name("Respuesta Método")));
        buttonElement.Click();
    }
    else
    {
        IWebElement buttonElement = wait.Until(ExpectedConditions.ElementToBeClickable(By.Name("Parámetros Método")));
        buttonElement.Click();
    }
}

            public static void CrudDetalle3(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string file)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TKactusDBgrid");
            if (tipo == 1)
            {
                Thread.Sleep(2000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                for (int i = 0; i < 2; i++)
                {
                    desktopSession.Keyboard.SendKeys(data[i]);
                    Thread.Sleep(500);
                    if (i == 1) { desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter); break; }
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                }
            }
            else
            {
                Thread.Sleep(2000);
                desktopSession.Keyboard.SendKeys(data[2]);
            }
        }
    }
}
