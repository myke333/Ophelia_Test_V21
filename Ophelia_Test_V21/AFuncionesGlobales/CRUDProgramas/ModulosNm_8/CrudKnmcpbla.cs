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
    class CrudKnmcpbla : FuncionesVitales
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



        static public void CrudPrincipal(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");
            //TPanel
            switch (bandera)
            {
                case 0:
                    //Marcar casillas
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 575, 20);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1500);
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 54, 78);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1500);
                   
                    break;
                case 1:
                    //INSERCION DE DATOS
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 51, 156);
                    desktopSession.Mouse.Click(null);
                    desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    Thread.Sleep(1500);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                    Thread.Sleep(1500);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    Thread.Sleep(1500);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    desktopSession.Keyboard.SendKeys(crudPrincipal[2]);
                    break;
            }
        }

        static public void botones(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");
            //TPanel
            switch (bandera)
            {
                case 0:
                    //Boton Datos
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 609, 226);
                    desktopSession.Mouse.Click(null);
                    break;
                case 1:
                    //Boton Aceptar
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 639, 445);
                    desktopSession.Mouse.Click(null);
                    break;
                case 2:
                    //Boton reportes
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 917, 445);
                    desktopSession.Mouse.Click(null);
                    break;
                case 3:
                    //Dar enter
                    WindowsDriver<WindowsElement> RootSession2 = null;
                    RootSession2 = PruebaCRUD.RootSession();
                    RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(2000);
                    break;
            }
        }






        }
}
