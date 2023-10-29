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
    class CrudKnmHenca : FuncionesVitales 
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

        static public void AgregarCrud(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal, string file)
        {

            var btnTDVI2 = desktopSession.FindElementsByClassName("TCaptura");
            var btnTDVI1 = desktopSession.FindElementsByClassName("TScrollBox");
            var btnTDVI = desktopSession.FindElementsByClassName("TTabSheet");
            //Debugger.Launch();
            //for (int i = 0; i < btnTDVI.Count; i++)
            //{
            //    btnTDVI[i].SendKeys(i.ToString());
            //}


            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 36, 67);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 140, 377);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 247, 230);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI1[0].Coordinates, 140, 10);
                desktopSession.Mouse.Click(null);
                newFunctions_4.ScreenshotNuevo("Ventana empleado", true, file);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI1[0].Coordinates, 153, 33);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI1[0].Coordinates, 405, 252);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[2]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI1[0].Coordinates, 217, 10);
                desktopSession.Mouse.Click(null);
                newFunctions_4.ScreenshotNuevo("Ventana empleado a reemplazar", true, file);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI1[0].Coordinates, 131, 31);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                newFunctions_4.ScreenshotNuevo("Mensaje de reemplazo", true, file);
                Thread.Sleep(1500);
                for (int a=0;a<4;a++)
                {
                    Enter(desktopSession);
                }
                //cODIGO USADO ANTERIORMENTE)
                //desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 36, 67);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1500);
                //desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 140, 377);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1500);
                //desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 290, 232);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1500);
                //WindowsDriver<WindowsElement> RootSession1 = null;
                //RootSession1 = PruebaCRUD.RootSession();
                //RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                //Thread.Sleep(2000);
                //desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 410, 232);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1500);
                //WindowsDriver<WindowsElement> RootSession2 = null;
                //RootSession2 = PruebaCRUD.RootSession();
                //RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                //Thread.Sleep(2000);
                //desktopSession.Mouse.MouseMove(btnTDVI1[0].Coordinates, 140, 10);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1500);
                //desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 600, 40);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1500);
                //WindowsDriver<WindowsElement> RootSession3 = null;
                //RootSession3 = PruebaCRUD.RootSession();
                //RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                //Thread.Sleep(1500);
                //WindowsDriver<WindowsElement> RootSession4 = null;
                //RootSession4 = PruebaCRUD.RootSession();
                //RootSession4.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                //Thread.Sleep(2000);
                //WindowsDriver<WindowsElement> RootSession5 = null;
                //RootSession5 = PruebaCRUD.RootSession();
                //RootSession5.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                //Thread.Sleep(2500);
                //desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 600, 70);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1500);
                //desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 400, 250);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1500);
                //desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                //Thread.Sleep(1000);
                //desktopSession.Mouse.MouseMove(btnTDVI1[0].Coordinates, 217, 10);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1500);
                //desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 580, 70);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(8000);
                //WindowsDriver<WindowsElement> RootSession6 = null;
                //RootSession6 = PruebaCRUD.RootSession();
                //RootSession6.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                //Thread.Sleep(30500);
                //WindowsDriver<WindowsElement> RootSession7 = null;
                //RootSession7 = PruebaCRUD.RootSession();
                //RootSession7.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                //Thread.Sleep(2000);
                //desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 390, 247);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1500);
                //desktopSession.Keyboard.SendKeys(crudPrincipal[2]);
                //Thread.Sleep(1000);

            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI1[0].Coordinates, 131, 31);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                newFunctions_4.ScreenshotNuevo("Mensaje de reemplazo", true, file);
                Thread.Sleep(1500);
                for (int a = 0; a < 4; a++)
                {
                    Enter(desktopSession);
                }

                //Codigo usado anteriormente
                //desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 390, 247);
                //desktopSession.Mouse.DoubleClick(null);
                //Thread.Sleep(1500);
                //desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                //Thread.Sleep(1000);

                //desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 101, 27);
                //desktopSession.Mouse.Click(null);
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                //desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
                //Thread.Sleep(1500);

            }
        }

        public static void EliminarRegistro(WindowsDriver<WindowsElement> desktopSession, WindowsElement ElementList, Dictionary<string, Point> botonesNavegador, string file)
        {
            desktopSession.Mouse.MouseMove(ElementList.Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(500);
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
            WindowsDriver<WindowsElement> RootSession2 = null;
            RootSession2 = PruebaCRUD.RootSession();
            RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2000);

        }


      


    }
}
