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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_8
{
    class CrudKnmdpcon : FuncionesVitales
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



        static public void InData(WindowsDriver<WindowsElement> desktopSession, string file, List<string> Principal)
        {
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");

            // Click forma de pago nomina.
            desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 493, 20);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 493, 39);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 493, 58);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 493, 71);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);


            // Click tipo de nomina normal
            desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 230, 71);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);


            // Click adicionales
            desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 448, 146);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 14, 160);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);



            // Click en Datos
            desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 590, 180);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);


            newFunctions_4.ScreenshotNuevo(" Programa parametrizado", true, file);

           
        }



        static public void ClickBtnAceptar(WindowsDriver<WindowsElement> desktopSession)
        {
            desktopSession.FindElementByName("Aceptar").Click();
            Thread.Sleep(1500);
        }



        

        static public void Aceptar1(WindowsDriver<WindowsElement> Session)
        {
            
            //Abre ventana export excel
            WindowsDriver<WindowsElement> RootSession2 = null;
            RootSession2 = PruebaCRUD.RootSession();

            RootSession2 = ReloadSession(RootSession2, "TFrmNmMsjInst");
            var btnTDVI2 = RootSession2.FindElementsByClassName("TFrmNmMsjInst");
            // Se escogen todos los >> atributos a exportar

            RootSession2.Mouse.MouseMove(btnTDVI2[0].Coordinates, 278, 134);
            RootSession2.Mouse.Click(null);
            Thread.Sleep(1000);

        }


        static public void Aceptar2(WindowsDriver<WindowsElement> Session)
        {
            
            //Abre ventana export excel
            WindowsDriver<WindowsElement> RootSession2 = null;
            RootSession2 = PruebaCRUD.RootSession();

            RootSession2 = ReloadSession(RootSession2, "TFrmMensaje");
            var btnTDVI2 = RootSession2.FindElementsByClassName("TFrmMensaje");

            // Se escogen todos los >> atributos a exportar
            RootSession2.Mouse.MouseMove(btnTDVI2[0].Coordinates, 39, 389);
            RootSession2.Mouse.Click(null);
            Thread.Sleep(1000);

            RootSession2.FindElementByName("Aceptar").Click();
            Thread.Sleep(6000);

            RootSession2.Mouse.Click(null);
            Thread.Sleep(2000);

        }




        static public void Enter(WindowsDriver<WindowsElement> desktopSession)
        {
            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

        }


        static public void CerrarVentana(WindowsDriver<WindowsElement> desktopSession)
        {
            // Se cierra ventana del proceso
            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();

            RootSession.FindElementByName("Close").Click();
            Thread.Sleep(1000);

        }
    }
}
