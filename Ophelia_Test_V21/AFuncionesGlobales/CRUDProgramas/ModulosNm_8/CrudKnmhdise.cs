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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_8
{
    class CrudKnmhdise : FuncionesVitales
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

        static public WindowsDriver<WindowsElement> ReloadSession1(WindowsDriver<WindowsElement> session, string parametro)
        {
            try
            {
                var ApplicationWindow = session.FindElementByClassName(parametro);
                var ApplicationSessionHandle = ApplicationWindow.GetAttribute("NativeWindowHandle");
                ApplicationSessionHandle = (int.Parse(ApplicationSessionHandle)).ToString("x"); //Convert to hex


                var appCapabilities = new AppiumOptions();
                appCapabilities.AddAdditionalCapability("appTopLevelWindow", ApplicationSessionHandle);
                var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);

                string ses = ApplicationSession.PageSource;
                return ApplicationSession;
            }
            catch (Exception ex) { return null; }
        }

        static public void AgregarCrud(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal)
        {

            //var btnTDVI = desktopSession.FindElementsByClassName("TDBEdit");
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");


            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 335, 40);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                WindowsDriver<WindowsElement> RootSession = null;
                RootSession = PruebaCRUD.RootSession();
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 600, 95);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                WindowsDriver<WindowsElement> RootSession1 = null;
                RootSession1 = PruebaCRUD.RootSession();
                RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);
                WindowsDriver<WindowsElement> RootSession2 = null;
                RootSession2 = PruebaCRUD.RootSession();
                RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);
                WindowsDriver<WindowsElement> RootSession3 = null;
                RootSession3 = PruebaCRUD.RootSession();
                RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 500, 145);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                WindowsDriver<WindowsElement> RootSession4 = null;
                RootSession4 = PruebaCRUD.RootSession();
                RootSession4.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 500, 185);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                WindowsDriver<WindowsElement> RootSession5 = null;
                RootSession5 = PruebaCRUD.RootSession();
                RootSession5.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 500, 225);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                WindowsDriver<WindowsElement> RootSession6 = null;
                RootSession6 = PruebaCRUD.RootSession();
                RootSession6.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 560, 147);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                desktopSession.Keyboard.SendKeys(crudPrincipal[0]);//100
                Thread.Sleep(1500);

            }
            else
            {

                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 180, 100);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(2000);
                desktopSession.Keyboard.SendKeys(crudPrincipal[1]);//3
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);

            }
        }

       

        static public void Enter(WindowsDriver<WindowsElement> desktopSession)
        {
            WindowsDriver<WindowsElement> RootSession2 = null;
            RootSession2 = PruebaCRUD.RootSession();
            RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2000);

        }

    }
}
