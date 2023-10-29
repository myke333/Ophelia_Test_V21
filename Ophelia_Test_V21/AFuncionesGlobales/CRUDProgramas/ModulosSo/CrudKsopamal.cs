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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSo
{
    class CrudKsopamal : FuncionesVitales
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
            //var btnTDVI = desktopSession.FindElementsByClassName("TTabSheet");
            //Debugger.Launch();
            //for (int i = 0; i < btnTDVI.Count; i++)
            //{
            //    btnTDVI[i].SendKeys(i.ToString());
            //}


            if (bandera == 0)
            {

                
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 200, 40);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                Thread.Sleep(1000);
                WindowsDriver<WindowsElement> RootSession = null;
                RootSession = PruebaCRUD.RootSession();
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 385, 40);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                WindowsDriver<WindowsElement> RootSession1 = null;
                RootSession1 = PruebaCRUD.RootSession();
                RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                

            }
            else
            {

                
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 345, 40);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                Thread.Sleep(1000);

            }
            
            if (bandera==3)
            {


                WindowsDriver<WindowsElement> RootSession26 = null;
                RootSession26 = PruebaCRUD.RootSession();
                RootSession26.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);
            }
        }
        static public void AgregarCrud1(WindowsDriver<WindowsElement> desktopSession, List<string> crudDet1)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

          
            {

                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 410, 170);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 150, 140);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDet1[0]);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 485, 140);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDet1[1]);
                Thread.Sleep(5000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 488, 170);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);

            }
        }

        static public void AgregarCrud2(WindowsDriver<WindowsElement> desktopSession, List<string> crudDet2)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            {
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 402, 355);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 150, 240);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDet2[0]);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 440, 240);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDet2[1]);
                Thread.Sleep(1000);
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 488, 355);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);

            }
        }

        static public void AgregarCrud3(WindowsDriver<WindowsElement> desktopSession, List<string> crudDet3)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 600, 143);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDet3[0]);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 815, 143);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDet3[1]);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1166, 143);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDet3[2]);
                Thread.Sleep(1000);

            }
        }

        public static void EliminarDetalle2(WindowsDriver<WindowsElement> desktopSession, string file)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");


            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 430, 355);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1500);
            newFunctions_4.ScreenshotNuevo("Borrar Registro", true, file);
            WindowsDriver<WindowsElement> RootSession1 = null;
            RootSession1 = PruebaCRUD.RootSession();
            RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            WindowsDriver<WindowsElement> RootSession2 = null;
            RootSession2 = PruebaCRUD.RootSession();
            RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            newFunctions_4.ScreenshotNuevo("Registro Borrado", true, file);
            Thread.Sleep(2000);
        }


        public static void EliminarDetalle1(WindowsDriver<WindowsElement> desktopSession, string file)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 444, 170);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1500);
            newFunctions_4.ScreenshotNuevo("Borrar Registro", true, file);
            WindowsDriver<WindowsElement> RootSession1 = null;
            RootSession1 = PruebaCRUD.RootSession();
            RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            WindowsDriver<WindowsElement> RootSession2 = null;
            RootSession2 = PruebaCRUD.RootSession();
            RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            newFunctions_4.ScreenshotNuevo("Registro Borrado", true, file);
            Thread.Sleep(2000);
        }


        static public void Enter(WindowsDriver<WindowsElement> desktopSession)
        {
            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2000);

        }

    }
}

