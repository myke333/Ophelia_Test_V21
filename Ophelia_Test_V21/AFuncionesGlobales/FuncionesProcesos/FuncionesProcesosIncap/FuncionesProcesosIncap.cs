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

namespace Ophelia_Test_V21.AFuncionesGlobales.FuncionesProcesos.FuncionesProcesosIncap
{
    class FuncionesProcesosIncap : FuncionesVitales
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

        static public void AgregarRegistro(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TDBEdit");
            //var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");
            var btnTDVI1 = desktopSession.FindElementByName("  Valores  ");
            //Debugger.Launch();
            //for (int i = 0; i < btnTDVI.Count; i++)
            //{
            //    btnTDVI[i].SendKeys(i.ToString());
            //    Thread.Sleep(10);
            //}


            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVI[14].Coordinates, 35, 15);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[16].Coordinates, 17, 4);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
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
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudPrincipal[4]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                DateTime fechaActual = DateTime.Today;
                string año = fechaActual.ToString("yyyy");
                string fecha = crudPrincipal[5]+ "/" + crudPrincipal[6] + "/" + año;
                desktopSession.Keyboard.SendKeys(fecha);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                

                //desktopSession.Mouse.MouseMove(btnTDVI[24].Coordinates, 27, 9);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1500);
                //desktopSession.Keyboard.SendKeys(crudPrincipal[5]);
                //Thread.Sleep(1000);
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                //Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI1.Coordinates, 110, 41);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI1.Coordinates, 238, 39);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);

            }

        }

        static public void CorrerNomina(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal)
        {

            var btnTDVI2 = desktopSession.FindElementByName("Aceptar");
            var btnTDVI = desktopSession.FindElementsByClassName("TEdit");
            var btnTDVI1 = desktopSession.FindElementByName("Simulación");
            //Debugger.Launch();
            //for (int i = 0; i < btnTDVI.Count; i++)
            //{
            //    btnTDVI[i].SendKeys(i.ToString());
            //    Thread.Sleep(10);
            //}


            if (bandera == 0)
            {
                //Primero obtenemos el día actual
                DateTime date = DateTime.Now;

                //Asi obtenemos el primer dia del mes actual
                DateTime oPrimerDiaDelMes = new DateTime(date.Year, date.Month, 1);

                //Y de la siguiente forma obtenemos el ultimo dia del mes
                //agregamos 1 mes al objeto anterior y restamos 1 día.
                DateTime UltimoDiaDelMes = oPrimerDiaDelMes.AddMonths(1).AddDays(-1);

                string Findate = UltimoDiaDelMes.ToString("dd/MM/yyyy");

                //desktopSession.Mouse.MouseMove(btnTDVI1.Coordinates, 3, 3);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[1].Coordinates, 64, 14);
                Thread.Sleep(1500);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                btnTDVI[1].Clear();
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(Findate);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(Findate);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                

            }
            else 
            {
                desktopSession.Mouse.MouseMove(btnTDVI2.Coordinates, 53, 14);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                
            }
        }

        static public void VerificarPrenomina(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal)
        {

            var btnTDVI = desktopSession.FindElementByName("GridPanel1");
            var btnTDVI1 = desktopSession.FindElementByClassName("TKNavegador");



            if (bandera == 0)
            {

                desktopSession.Mouse.MouseMove(btnTDVI.Coordinates, 437, 22);
                Thread.Sleep(1500);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
               
            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI1.Coordinates, 69, 12);
                Thread.Sleep(1500);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
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
