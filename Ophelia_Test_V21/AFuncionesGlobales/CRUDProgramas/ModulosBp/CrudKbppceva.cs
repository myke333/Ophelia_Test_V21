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

using System.Windows.Forms;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBp
{
    class CrudKbppceva : FuncionesVitales
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
            var btnTDVI = desktopSession.FindElementsByClassName("TTabSheet");
            //var btnTDVI = desktopSession.FindElementsByClassName("TTabSheet");
            //Debugger.Launch();
            //for (int i = 0; i < btnTDVI.Count; i++)
            //{
            //    btnTDVI[i].SendKeys(i.ToString());
            //}


            if (bandera == 0)
            {
                //codigo centro vacacional 
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 44, 29);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                Thread.Sleep(1000);

                //nombre centro vacacional 
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 237, 29);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                Thread.Sleep(1000);

                //cabañas 
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 825, 29);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[12]);
                Thread.Sleep(1000);

                //pais 
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 27, 76);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[2]);
                Thread.Sleep(1000);

                //departamento
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 292, 76);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
                Thread.Sleep(1000);

                //municipio
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 552, 76);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[4]);
                Thread.Sleep(1000);

                //tiempo solicitud 
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 47, 116);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[5]);
                Thread.Sleep(1000);


                //cantidad personas 
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 167, 116);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[6]);
                Thread.Sleep(1000);


                //numero noches
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 274, 116);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[7]);
                Thread.Sleep(1000);


                //adulto 
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 390, 116);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[8]);
                Thread.Sleep(1000);


                //cantidad niños
             
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 519, 116);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[9]);
                Thread.Sleep(1000);


                //años nueva  
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 630, 116);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[10]);
                Thread.Sleep(1000);


               


                //concepto nomina 
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 371, 188);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[11]);
                Thread.Sleep(1000);

               


            }
            else
            {
                //años nueva  
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 630, 116);
                desktopSession.Mouse.DoubleClick(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[10]);
                Thread.Sleep(1000);

                //desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 101, 27);
                //desktopSession.Mouse.Click(null);
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                //desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
                //Thread.Sleep(1500);

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