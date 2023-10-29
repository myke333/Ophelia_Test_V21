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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_9
{
    class CrudKnmpreai : FuncionesVitales
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

        static public void AgregarCrud(WindowsDriver<WindowsElement> desktopSession)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TTabSheet");

            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 123, 43);
            desktopSession.Mouse.Click(null);
            desktopSession.Keyboard.SendKeys("01/01/2021");
            Thread.Sleep(2000);
                        
            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 264, 43);
            desktopSession.Mouse.Click(null);
            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
            desktopSession.Keyboard.SendKeys("2021");
            Thread.Sleep(2000);

            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 362, 43);
            desktopSession.Mouse.Click(null);
            desktopSession.Keyboard.SendKeys("1");
            Thread.Sleep(2000);

            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 487, 43);
            desktopSession.Mouse.Click(null);
            desktopSession.Keyboard.SendKeys("01/01/2021");
            Thread.Sleep(2000);

            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 430, 172);
            desktopSession.Mouse.Click(null);
            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 430, 193);
            desktopSession.Mouse.Click(null);

            desktopSession.FindElementByName("Aceptar").Click();

        }

        public static void AgregarCrud2(WindowsDriver<WindowsElement> desktopSession)
        {
            Thread.Sleep(3000);
            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(3000);

        }

        public static void AgregarCrud3(WindowsDriver<WindowsElement> desktopSession)
        {
            desktopSession.FindElementByName("Actualizar").Click();
            Thread.Sleep(10000);
        }

        public static void AgregarCrud4(WindowsDriver<WindowsElement> desktopSession)
        {
            Thread.Sleep(3000);
            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(3000);

        }
    }
}
