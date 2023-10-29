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

namespace Ophelia_Test_V21.TestMethods.ModulosNM
{
    class crudKnmdgapo : FuncionesVitales
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

            var btnTDVI = desktopSession.FindElementsByClassName("TGroupBox");
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");
            //var btnTDVI = desktopSession.FindElementsByClassName("TPanel");
            //Debugger.Launch();

            desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 118, 31);
            desktopSession.Mouse.DoubleClick(null);
            desktopSession.Keyboard.SendKeys("03/03/2020");
            Thread.Sleep(2000);

            desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 71, 70);
            desktopSession.Mouse.Click(null);
            desktopSession.Keyboard.SendKeys("5000");
            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            Thread.Sleep(2000);

            desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 382, 70);
            desktopSession.Mouse.Click(null);
            desktopSession.Keyboard.SendKeys("1");
            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            Thread.Sleep(2000);

        }

        static public void Boton(WindowsDriver<WindowsElement> desktopSession)
        {
            desktopSession.FindElementByName("Aceptar").Click();

        }

        static public void Enter(WindowsDriver<WindowsElement> desktopSession)
        {

            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

        }

    }
}

