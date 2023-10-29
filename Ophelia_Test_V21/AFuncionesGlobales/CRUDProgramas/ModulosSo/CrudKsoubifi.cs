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
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosGn
{
    class CrudKsoubifi /*: FuncionesVitales*/
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



        static public void AggCrudPri(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal)
        {

            var btnTDVI0 = desktopSession.FindElementsByClassName("TScrollBox");


            if (bandera == 0)
            {
                //Escoge Dependencia.
                desktopSession.Mouse.MouseMove(btnTDVI0[0].Coordinates, 918, 87);
                desktopSession.Mouse.Click(null);

                WindowsDriver<WindowsElement> RootSession = null;
                RootSession = PruebaCRUD.RootSession();
                RootSession = ReloadSession(RootSession, "TFrmGnArbol");

                var btnTDVI = RootSession.FindElementsByClassName("TVirtualStringTree");

                RootSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 109, 8);
                RootSession.Mouse.DoubleClick(null);
                Thread.Sleep(1200);
                RootSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 109, 26);
                RootSession.Mouse.DoubleClick(null);
                Thread.Sleep(1200);
                RootSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 97, 45);
                RootSession.Mouse.DoubleClick(null);
                Thread.Sleep(1200);



                //click btn aceptar en la ventana dependencia
                //WindowsDriver<WindowsElement> RootSession1 = null;
                //RootSession = PruebaCRUD.RootSession();
                //RootSession = ReloadSession(RootSession, "TFrmGnArbol");

                //var btnTDVI1 = RootSession.FindElementsByClassName("TBitBtn");
                desktopSession.Mouse.MouseMove(btnTDVI0[0].Coordinates, 767, 413);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1200);
                //Escoge Dependencia.


                // Click Zona.
                desktopSession.Mouse.MouseMove(btnTDVI0[0].Coordinates, 57, 134);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                Thread.Sleep(1000);
                
                // ingresa datos centro de costos
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(900);
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(900);
                desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                Thread.Sleep(1500); 
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(900);

                // Click en Pais.
                //desktopSession.Mouse.MouseMove(btnTDVI0[0].Coordinates, 50, 174);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(750);
                //desktopSession.Keyboard.SendKeys(crudPrincipal[4]);
                //Thread.Sleep(750);
                //RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                //Thread.Sleep(900);
                //RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                //Thread.Sleep(900);
                //// ingresa datos en departamentos
                //desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
                //Thread.Sleep(1000);

                desktopSession.Mouse.MouseMove(btnTDVI0[0].Coordinates, 280, 169);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1200);

                desktopSession.Mouse.MouseMove(btnTDVI0[0].Coordinates, 661, 142);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1200);

                desktopSession.Mouse.MouseMove(btnTDVI0[0].Coordinates, 682, 34);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1200);

                desktopSession.Mouse.MouseMove(btnTDVI0[0].Coordinates, 696, 63);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1200);
            }
            else
            {
                // Click Zona para Actdatos.
                desktopSession.Mouse.MouseMove(btnTDVI0[0].Coordinates, 57, 134);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(750);
                // Act datos centro de costos
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(900);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(900);
                desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                Thread.Sleep(1000);

            }
        }

        static public void Qbe2(WindowsDriver<WindowsElement> desktopSession, string file)
        {
            var btnTDVI2 = desktopSession.FindElementsByClassName("TDBGrid");

            desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 1292, 245);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1500);
            newFunctions_4.ScreenshotNuevo("Click QBE Det 1", true, file);


            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession = ReloadSession(RootSession, "TKQBE");

            var btnTDVI = RootSession.FindElementsByClassName("TTabSheet");

            RootSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 600, 16);
            RootSession.Mouse.Click(null);
            Thread.Sleep(2000);

            WindowsDriver<WindowsElement> RootSession2 = null;
            RootSession2 = PruebaCRUD.RootSession();
            RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2000);
            newFunctions_4.ScreenshotNuevo("Registros QBE", true, file);

            WindowsDriver<WindowsElement> RootSession3 = null;
            RootSession3 = PruebaCRUD.RootSession();
            RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2000);
        }

        static public void AggCrudDet(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDet1)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TDBGrid");

            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 37, 30);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(800);
                desktopSession.Keyboard.SendKeys(crudDet1[0]);
                Thread.Sleep(800);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(900);
                desktopSession.Keyboard.SendKeys(crudDet1[1]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(800);

            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 37, 30);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDet1[2]);
                Thread.Sleep(1000);

            }
        }

    }
}