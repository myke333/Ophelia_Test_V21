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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_11
{
    class CrudKnmdpcol : FuncionesVitales
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


        static public void AgregarCrud(WindowsDriver<WindowsElement> desktopSession, List<string> crudPrincipal)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

           

                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 95,53);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                WindowsDriver<WindowsElement> RootSession = null;
                RootSession = PruebaCRUD.RootSession();
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 205, 53);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                WindowsDriver<WindowsElement> RootSession1 = null;
                RootSession1 = PruebaCRUD.RootSession();
                RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);
            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 50, 92);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1500);
            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 315, 50);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1500);
            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 315, 90);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1500);
            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 530, 56);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1500);
            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 200, 210);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1500);
            desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
            Thread.Sleep(1500);
            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 130, 170);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1500);
            //WindowsDriver<WindowsElement> RootSession = null;
            //RootSession = PruebaCRUD.RootSession();
            //RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
        }

        static public void Enter(WindowsDriver<WindowsElement> desktopSession)
        {

            WindowsDriver<WindowsElement> RootSession2 = null;
            RootSession2 = PruebaCRUD.RootSession();
            RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

        }

        static public void Aceptar(WindowsDriver<WindowsElement> desktopSession)

        {
            desktopSession.FindElementByName("Aceptar").Click();



        }

    }
}
