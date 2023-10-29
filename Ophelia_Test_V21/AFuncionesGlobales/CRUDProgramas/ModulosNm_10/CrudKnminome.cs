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
    class CrudKnminome : FuncionesVitales
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

        static public void ClickButtonAceptarEnter(WindowsDriver<WindowsElement> desktopSession, int bandera)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TBitBtn");
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");
            //var btnTDVI = desktopSession.FindElementsByClassName("TPanel");
            //Debugger.Launch();


            if (bandera == 0)
            {
                var ElementList = desktopSession.FindElementsByName("Aceptar");
                ElementList[0].Click();
                Thread.Sleep(3000);


            }
            else
            {
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1500);
                WindowsDriver<WindowsElement> RootSession = null;
                RootSession = PruebaCRUD.RootSession();
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                //Thread.Sleep(1500);

                //desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 101, 27);
                //desktopSession.Mouse.Click(null);
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                //desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
                //Thread.Sleep(1500);

            }
        }


        static public void Click(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal, string file)
        {

            //var btnTDVI2 = desktopSession.FindElementsByClassName("TDBEdit");
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");            
            //Debugger.Launch();

            if (bandera == 0)
            {
                
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 141, 106);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 347, 213);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                newFunctions_4.ScreenshotNuevo("Datos de configuracion de conexiones", true, file);
                Thread.Sleep(1000);
                ////////////////////////////////
                WindowsDriver<WindowsElement> RootSession = null;
                RootSession = PruebaCRUD.RootSession();
                RootSession = ReloadSession(RootSession, "TfrmConfig");

                var btnTDVI = RootSession.FindElementsByClassName("TTabSheet");

                RootSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 318, -54);
                RootSession.Mouse.Click(null);
                Thread.Sleep(2000);
                WindowsDriver<WindowsElement> RootSession2 = null;
                RootSession2 = PruebaCRUD.RootSession();
                RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);
                newFunctions_4.ScreenshotNuevo("Se elimina el registro", true, file);
                Thread.Sleep(1000);

                RootSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 147, 70);
                RootSession.Mouse.Click(null);
                Thread.Sleep(1000);
                RootSession.Keyboard.SendKeys(crudPrincipal[1]);
                Thread.Sleep(1500);
                RootSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 435, 70);
                RootSession.Mouse.Click(null);
                Thread.Sleep(1000);
                RootSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 435, 100);
                RootSession.Mouse.Click(null);
                Thread.Sleep(1000);
                RootSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 213, 157);
                RootSession.Mouse.Click(null);
                Thread.Sleep(1000);
                RootSession.Mouse.Click(null);
                Thread.Sleep(1000);
                RootSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 213, 170);
                RootSession.Mouse.Click(null);
                Thread.Sleep(1000);
                RootSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 402, 157);
                RootSession.Mouse.Click(null);
                Thread.Sleep(1000);
                RootSession.Keyboard.SendKeys(crudPrincipal[2]);
                Thread.Sleep(1500);
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);

                RootSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 213, 171);
                RootSession.Mouse.Click(null);
                Thread.Sleep(1000);
                RootSession.Mouse.Click(null);
                Thread.Sleep(1000);
                RootSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 213, 198);
                RootSession.Mouse.Click(null);
                Thread.Sleep(1000);
                RootSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 402, 171);
                RootSession.Mouse.Click(null);
                Thread.Sleep(1000);
                RootSession.Keyboard.SendKeys(crudPrincipal[3]);
                Thread.Sleep(1500);
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);

                RootSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 213, 185);
                RootSession.Mouse.Click(null);
                Thread.Sleep(1000);
                RootSession.Mouse.Click(null);
                Thread.Sleep(1000);
                RootSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 213, 235);
                RootSession.Mouse.Click(null);
                Thread.Sleep(1000);
                RootSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 402, 185);
                RootSession.Mouse.Click(null);
                Thread.Sleep(1000);
                RootSession.Keyboard.SendKeys(crudPrincipal[4]);
                Thread.Sleep(1500);

                RootSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 112, 325);
                RootSession.Mouse.Click(null);
                Thread.Sleep(1000);
                RootSession.Keyboard.SendKeys(crudPrincipal[5]);
                Thread.Sleep(1500);
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                RootSession.Keyboard.SendKeys(crudPrincipal[6]);
                Thread.Sleep(1500);

                RootSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 273, 532);
                RootSession.Mouse.Click(null);
                Thread.Sleep(1000);
                RootSession.Keyboard.SendKeys(crudPrincipal[7]);
                Thread.Sleep(1500);
                newFunctions_4.ScreenshotNuevo("Datos ingresados", true, file);
                //RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                //Thread.Sleep(1000);
                RootSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 708, 54);
                RootSession.Mouse.Click(null);
                Thread.Sleep(1000);
                newFunctions_4.ScreenshotNuevo("Se comprueba la conexion", true, file);
                Thread.Sleep(7000);
                newFunctions_4.ScreenshotNuevo("Conexion exitosa", true, file);
                Thread.Sleep(1000);                
                RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);

                RootSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 260, -54);
                RootSession.Mouse.Click(null);
                Thread.Sleep(2000);
                RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);
                newFunctions_4.ScreenshotNuevo("Se guarda el registro", true, file);
                Thread.Sleep(1000);

                RootSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 830, -90);
                RootSession.Mouse.Click(null);
                Thread.Sleep(2000);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 100, 11);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);

            }
            if (bandera == 1)
            {
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 900, 383);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                //WindowsDriver<WindowsElement> RootSession = null;
                //RootSession = PruebaCRUD.RootSession();
                //RootSession.Keyboard.SendKeys("ReportePDF1_Preliminar_08022022_161632");
                //Thread.Sleep(1000);
                //RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

            }
            if (bandera == 2)
            {
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 430, -30);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);
                //WindowsDriver<WindowsElement> RootSession = null;
                //RootSession = PruebaCRUD.RootSession();
                //RootSession.Keyboard.SendKeys("ReportePDF1_Preliminar_08022022_161632");
                //Thread.Sleep(1000);
                //RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

            }
        }

        public static void ClickButtonErrores(WindowsDriver<WindowsElement> desktopSession, string file)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByName("Errores");
            newFunctions_4.ScreenshotNuevo("Errores", true, file);
            ElementList[0].Click();
        }


    }
}