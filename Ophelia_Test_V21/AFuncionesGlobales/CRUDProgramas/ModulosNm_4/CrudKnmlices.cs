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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_4
{

    class CrudKnmlices : FuncionesVitales
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




        public static void CrudKNmlices(WindowsDriver<WindowsElement> desktopSession, string file)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");
            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 676, 441);
            desktopSession.Mouse.Click(null);
            newFunctions_4.ScreenshotNuevo("Boton de aceptar", true, file);
            Thread.Sleep(2000);
            Enter(desktopSession);
            newFunctions_4.ScreenshotNuevo("Proceso cargado", true, file);
            Thread.Sleep(2000);
            Enter(desktopSession);
            newFunctions_4.ScreenshotNuevo("Mensaje de error", true, file);
            Thread.Sleep(2000);
            Enter(desktopSession);
            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 889, 439);
            desktopSession.Mouse.Click(null);
            newFunctions_4.ScreenshotNuevo("Reporte de errores", true, file);

        }

        public static void Enter(WindowsDriver<WindowsElement> desktopSession)
        {
        WindowsDriver<WindowsElement> RootSession2 = null;
        RootSession2 = PruebaCRUD.RootSession();
        RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
        Thread.Sleep(2000);
        }


      

    }

}
