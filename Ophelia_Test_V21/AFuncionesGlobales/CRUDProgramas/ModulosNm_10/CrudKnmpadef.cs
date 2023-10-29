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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_10
{
    class CrudKnmpadef : FuncionesVitales
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

        static public void AgregarCrud(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal, string file)
        {

            //var btnTDVI = desktopSession.FindElementsByClassName("TDBEdit");
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");
            //var btnTDVI = desktopSession.FindElementsByClassName("TTabSheet");
            //Debugger.Launch();
            //for (int i = 0; i < btnTDVI.Count; i++)
            //{
            //    btnTDVI[i].SendKeys(i.ToString());
            //}


            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 498, 333);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudPrincipal[6]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 20, 40);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 20, 60);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 20, 80);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 20, 120);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 205, 40);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 205, 60);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 30, 152);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 456, 86);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 650, 86);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 458, 50);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 113, 4);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 40, 271);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudPrincipal[2]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 456, 252);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 444, 441);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                newFunctions_4.ScreenshotNuevo("Datos pestana Informacion 2", true, file);
                Thread.Sleep(3000);

                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 210, 10);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 25, 56);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                newFunctions_4.ScreenshotNuevo("Datos pestana Reportes", true, file);
                Thread.Sleep(3000);

                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 30, 10);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 166, 65);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 166, 98);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 166, 129);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 322, 80);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[4]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudPrincipal[5]);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 50, 240);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                //desktopSession.Keyboard.SendKeys(crudPrincipal[6]);
                //Thread.Sleep(1000);
                //desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 50, 230);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[7]);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 352, 160);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 600, 190);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 70, 290);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 187, 290);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[8]);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 50, 356);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[9]);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 50, 387);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[10]);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 570, 356);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                newFunctions_4.ScreenshotNuevo("Datos pestana Informacion 1", true, file);
                Thread.Sleep(3000);

            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 50, 387);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[11]);
                Thread.Sleep(1000);

                //desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 101, 27);
                //desktopSession.Mouse.Click(null);
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                //desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
                //Thread.Sleep(1500);

            }
        }

        static public void Click(WindowsDriver<WindowsElement> desktopSession, int bandera)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 396, -20);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);

            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 46, 222);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 90, 272);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1000);

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