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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosAc
{
    class CrudKacplanh
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



        static public void AggDataCrudPrincial(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            if (bandera == 0)  // agg data
            {

                // lupa 1
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 606, 41);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);


                WindowsDriver<WindowsElement> RootSession = null;
                RootSession = PruebaCRUD.RootSession();
                RootSession = ReloadSession(RootSession, "TKQBE");
                
                var btnTDVI2 = RootSession.FindElementsByClassName("TTabSheet");
                Thread.Sleep(500);

                // click en aceptar para traer todos los registros
                RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 592, 24);
                RootSession.Mouse.Click(null);
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);




                // Ingresa el dato en la ventana normal de la lupa
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

                RootSession1.Keyboard.SendKeys(crudPrincipal[0]);
                Thread.Sleep(500);

                RootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1500);

                ///////////////////////////////////////////////////////////////


                /////////////////   Lupa 2
                ///
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 606, 87);
                desktopSession.Mouse.Click(null);
                WindowsDriver<WindowsElement> RootSession2 = null;
                RootSession2 = PruebaCRUD.RootSession();

                //RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                //Thread.Sleep(500);
                //RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                //Thread.Sleep(500);
                //RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                //Thread.Sleep(1000);
                
                RootSession2.Keyboard.SendKeys(crudPrincipal[1]);
                Thread.Sleep(1000);


                RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

                Thread.Sleep(1500);

                ///////////////////////////

                ///////////   Lupa 3 
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 606, 127);
                desktopSession.Mouse.Click(null);
                RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1500);

                //////////////////////////////

                ////////////    fecha de corte
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 196, 212);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
                Thread.Sleep(1500);

                ///////////////////////////////

                //Tipo de cargo - Lupa 4
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 718, 399);
                desktopSession.Mouse.Click(null);
                WindowsDriver<WindowsElement> RootSession4 = null;
                RootSession4 = PruebaCRUD.RootSession();


                RootSession4.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                RootSession4.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                RootSession4.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                RootSession4.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1500);

                RootSession4.Keyboard.SendKeys(crudPrincipal[5]);
                Thread.Sleep(500);

                RootSession4.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1500);

            }
            else
            {
                ///////////  Actualizar   fecha de corte
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 196, 212);
                desktopSession.Mouse.DoubleClick(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[4]);
                Thread.Sleep(500);
                ///////////////////////////////

            }
        }


        static public void BtnAceptar(WindowsDriver<WindowsElement> desktopSession, string file)
        {
            desktopSession.FindElementByName("Aceptar").Click();
            Thread.Sleep(2800);
            newFunctions_4.ScreenshotNuevo("Final del proceso", true, file);
        }


        static public void Enter(WindowsDriver<WindowsElement> desktopSession)
        {            
            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
        }





    }
}
