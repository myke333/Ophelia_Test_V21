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


namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosGe
{
    class CrudKgepbsca
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




        static public void AgregarCrud(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");


            if (bandera == 0)
            {

                //---------ingresa Tipo Indicador
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 81, 36);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(500);



                // --------ingresa Codigo
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 204, 36);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                Thread.Sleep(1000);


                // --------ingresa NombreIndicador
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[2]);
                Thread.Sleep(1000);


                // -------ingresa Pespectiva en Lupa1 
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 600, 140);
                desktopSession.Mouse.Click(null);


                WindowsDriver<WindowsElement> RootSession1 = null;
                RootSession1 = PruebaCRUD.RootSession();
                Thread.Sleep(1000);
                RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1500);

                RootSession1.Keyboard.SendKeys(crudPrincipal[4]);
                Thread.Sleep(500);

                RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1500);


                // -------ingresa Obj Estrategico en Lupa2 
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 600, 197 );
                desktopSession.Mouse.Click(null);
                Thread.Sleep(800);

                RootSession1.Keyboard.SendKeys(crudPrincipal[5]);
                Thread.Sleep(500);

                RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1500);


                // -------ingresa Proyecto en Lupa3 
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 600, 247);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(800);

                RootSession1.Keyboard.SendKeys(crudPrincipal[6]);
                Thread.Sleep(500);

                RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1500);


                // -------ingresa Area en Lupa4 
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 600, 292);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(800);

                RootSession1.Keyboard.SendKeys(crudPrincipal[7]);
                Thread.Sleep(500);

                RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1500);



                // -------ingresa Obje Are en Lupa5 
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 600, 343);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(800);

                RootSession1.Keyboard.SendKeys(crudPrincipal[8]);
                Thread.Sleep(500);

                RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1500);



                // -------ingresa Iniciativa Objetivo Area  en Lupa6 
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 600, 397);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(800);

                RootSession1.Keyboard.SendKeys(crudPrincipal[9]);
                Thread.Sleep(500);

                RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1500);




            }
            else
            {                
                // --------ingresa Codigo
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 204, 36);
                desktopSession.Mouse.Click(null);
                //desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                //Thread.Sleep(1000);


                // --------ingresa NombreIndicador
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
                Thread.Sleep(1000);

            }
        }


    }
}
