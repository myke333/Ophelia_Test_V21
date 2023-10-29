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

namespace Ophelia_Test_V21.AFuncionesGlobales.FuncionesProcesos.FuncionesProcesosKNmNoved
{
    class FuncionesProcesosKNmNoved
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

        static public void AgregarRegistro(WindowsDriver<WindowsElement> desktopSession, string bandera, List<string> crudPrincipal)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TDBEdit");
            

            if (bandera == "F")
            {
                //IDENTIFICACION
                desktopSession.Mouse.MouseMove(btnTDVI[16].Coordinates, 50, 8);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                Thread.Sleep(1000);
            

                //CONCEPTO
                desktopSession.Mouse.MouseMove(btnTDVI[3].Coordinates, 33, 6);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                Thread.Sleep(1000);
                

                //VALOR CUOTA
                desktopSession.Mouse.MouseMove(btnTDVI[11].Coordinates, 8, 13);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[4]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(5000);

                //NUMERO CUOTAS
                desktopSession.Mouse.MouseMove(btnTDVI[7].Coordinates, 39, 5);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Delete);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[5]);
                Thread.Sleep(1000);

                DateTime fechaActual = DateTime.Today;
                string año = fechaActual.ToString("yyyy");
                string fecha = crudPrincipal[6] + "/" + crudPrincipal[7] + "/" + año;

                //FECHA PAGO
                desktopSession.Mouse.MouseMove(btnTDVI[5].Coordinates, 4, 7);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Delete);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(fecha);
                Thread.Sleep(1000);

                //FECHA NOVEDAD
                desktopSession.Mouse.MouseMove(btnTDVI[6].Coordinates, 4, 10);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Delete);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(fecha);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1500);



            }

            if (bandera == "V")
            {
                //IDENTIFICACION
                desktopSession.Mouse.MouseMove(btnTDVI[16].Coordinates, 50, 8);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                Thread.Sleep(1000);


                //CONCEPTO
                desktopSession.Mouse.MouseMove(btnTDVI[3].Coordinates, 33, 6);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                Thread.Sleep(1000);


               
                //CANTIDAD
                desktopSession.Mouse.MouseMove(btnTDVI[10].Coordinates, 10, 13);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Delete);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
                Thread.Sleep(1000);


                DateTime fechaActual = DateTime.Today;
                string año = fechaActual.ToString("yyyy");
                string fecha = crudPrincipal[6] + "/" + crudPrincipal[7] + "/" + año;

                //FECHA PAGO
                desktopSession.Mouse.MouseMove(btnTDVI[5].Coordinates, 4, 7);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Delete);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(fecha);
                Thread.Sleep(1000);

                //FECHA NOVEDAD
                desktopSession.Mouse.MouseMove(btnTDVI[6].Coordinates, 4, 10);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Delete);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(fecha);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1500);



            }

        }
    }
}
