using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_5
{
    class CrudKnmadmpr
    {
        protected static WindowsDriver<WindowsElement> rootSession;
        public static WindowsDriver<WindowsElement> RootSession()
        {
            var appCapabilities = new AppiumOptions();
            appCapabilities.AddAdditionalCapability("app", "Root");
            var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
            return ApplicationSession;
        }
        static public WindowsDriver<WindowsElement> ReloadSession(WindowsDriver<WindowsElement> session, string Clase)
        {
            try
            {
                var ApplicationWindow = session.FindElementByClassName(Clase);
                var ApplicationSessionHandle = ApplicationWindow.GetAttribute("NativeWindowHandle");
                ApplicationSessionHandle = (int.Parse(ApplicationSessionHandle)).ToString("x"); //Convert to hex


                var appCapabilities = new AppiumOptions();
                //appCapabilities.AddAdditionalCapability("app", "Root");
                appCapabilities.AddAdditionalCapability("appTopLevelWindow", ApplicationSessionHandle);
                var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);

                string ses = ApplicationSession.PageSource;
                ////Debugger.Launch();
                return ApplicationSession;
            }
            catch (Exception ex) { return null; }
        }
        public static void CrudKNmadmpr(WindowsDriver<WindowsElement> desktopSession, string file)
        {
            desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Actualizar Tiempo Contrato").Coordinates);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(500);
            desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Aceptar").Coordinates);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(500);
            PruebaCRUD.cerrarBorrar(desktopSession);
            Thread.Sleep(1000);
            rootSession = RootSession();
            rootSession = ReloadSession(rootSession, "TFrmNmCapDa");
            Thread.Sleep(500);
            rootSession.Mouse.MouseMove(rootSession.FindElementByName("Aceptar").Coordinates);
            rootSession.Mouse.Click(null);
            Thread.Sleep(2000);
            rootSession = RootSession();
            var allFrame = rootSession.FindElementsByClassName("TFrmNmCapDa");
            new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset(allFrame[0].Size.Width / 2, allFrame[0].Size.Height / 2).DoubleClick().Click().Perform();
            rootSession = RootSession();
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2000);
            newFunctions_4.ScreenshotNuevo("Proceso Iniciado", true, file);
            PruebaCRUD.cerrarBorrar(desktopSession);
            Thread.Sleep(5000);
            newFunctions_4.ScreenshotNuevo("Proceso Terminado", true, file);
            PruebaCRUD.cerrarBorrar(desktopSession);
        }
    }
}
