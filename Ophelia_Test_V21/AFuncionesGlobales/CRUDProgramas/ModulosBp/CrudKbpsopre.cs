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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBp
{

    class CrudKbpsopre : FuncionesVitales
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

            var btnTDVI = desktopSession.FindElementsByClassName("TDBEdit");

            /*for (int i = 0; i < btnTDVI.Count; i++)
            {
                btnTDVI[i].SendKeys(i.ToString());
            }*/

            if (bandera == 0)
            {
                desktopSession.Keyboard.SendKeys(crudPrincipal[9]);
                Thread.Sleep(2500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(8000);
                WindowsDriver <WindowsElement> RootSession = null;
                RootSession = PruebaCRUD.RootSession();
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(8000);

            }
            else
            {

            }

        }

        static public void AgregarCrud2(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal)
        {

            //var btnTDVI = desktopSession.FindElementsByClassName("TDBEdit");
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");
            var btnTDVI1 = desktopSession.FindElementsByClassName("TTabSheet");
            //Debugger.Launch();
            //for (int i = 0; i < btnTDVI.Count; i++)
            //{
            //    btnTDVI[i].SendKeys(i.ToString());
            //}


            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 150, 100);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2500);
                desktopSession.Mouse.MouseMove(btnTDVI1[0].Coordinates, 157, 35);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                Thread.Sleep(2500);
                desktopSession.Mouse.MouseMove(btnTDVI1[0].Coordinates, 70, 84);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                Thread.Sleep(2500);
                desktopSession.Mouse.MouseMove(btnTDVI1[0].Coordinates, 195, 84);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[2]);
                Thread.Sleep(2500);
                desktopSession.Mouse.MouseMove(btnTDVI1[0].Coordinates, 40, 124);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
                Thread.Sleep(2500);
                desktopSession.Mouse.MouseMove(btnTDVI1[0].Coordinates, 40, 164);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[4]);
                Thread.Sleep(2500);
                desktopSession.Mouse.MouseMove(btnTDVI1[0].Coordinates, 195, 122);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[5]);
                Thread.Sleep(2500);
                desktopSession.Mouse.MouseMove(btnTDVI1[0].Coordinates, 70, 228);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[6]);
                Thread.Sleep(2500);
                desktopSession.Mouse.MouseMove(btnTDVI1[0].Coordinates, 180, 228);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[7]);
                Thread.Sleep(2500);
                desktopSession.Mouse.MouseMove(btnTDVI1[0].Coordinates, 621, 40);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2500);
                WindowsDriver<WindowsElement> RootSession = null;
                RootSession = PruebaCRUD.RootSession();
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(5000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 225, 100);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2500);
                desktopSession.Mouse.MouseMove(btnTDVI1[0].Coordinates, 133, 43);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2500);
                WindowsDriver<WindowsElement> RootSession1 = null;
                RootSession1 = PruebaCRUD.RootSession();
                RootSession1 = ReloadSession(RootSession1, "TFrmBpLiqui");
                var btnTDVI2 = RootSession1.FindElementsByClassName("TFrmBpLiqui");
                RootSession1.Mouse.MouseMove(btnTDVI2[0].Coordinates, 268, 310);
                RootSession1.Mouse.Click(null);
                Thread.Sleep(10000);
                WindowsDriver<WindowsElement> RootSession2 = null;
                RootSession2 = PruebaCRUD.RootSession();
                RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(18000);
                WindowsDriver<WindowsElement> RootSession3 = null;
                RootSession3 = PruebaCRUD.RootSession();
                RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(10000);
                WindowsDriver<WindowsElement> RootSession4 = null;
                RootSession4 = PruebaCRUD.RootSession();
                RootSession4.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(10000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 150, 100);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2500);

            }
            else
            {

                Thread.Sleep(2500);
                desktopSession.Mouse.MouseMove(btnTDVI1[0].Coordinates, 157, 35);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(2500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[8]);
                Thread.Sleep(2500);


            }
        }
    }
}