using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
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
    class CrudKnmacenc : FuncionesVitales
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
                appCapabilities.AddAdditionalCapability("appTopLevelWindow", ApplicationSessionHandle);
                var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);

                string ses = ApplicationSession.PageSource;
                return ApplicationSession;
            }
            catch (Exception ex) { return null; }
        }
        public static void CrudKNmacenc(WindowsDriver<WindowsElement> desktopSession, int Tipo, string file)
        {
            if(Tipo == 1)
            {
                string pdf = @"C:\Reportes\ReportePdf_" + "Previo" + "_" + Hora();
                desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Previo").Coordinates);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TFrmReporte");
                Thread.Sleep(500);
                rootSession.Mouse.MouseMove(rootSession.FindElementByName("Aceptar").Coordinates);
                rootSession.Mouse.Click(null);
                Thread.Sleep(500);
                newFunctions_5.GenerarReportes("Previo", file, pdf);
            }
            else
            {
                desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Aceptar").Coordinates);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                PruebaCRUD.cerrarBorrar(desktopSession);
                Thread.Sleep(10000);
                PruebaCRUD.cerrarBorrar(desktopSession);
                Thread.Sleep(5000);
                newFunctions_4.ScreenshotNuevo("Proceso Ejecutado", true, file);
                PruebaCRUD.cerrarBorrar(desktopSession);
                Thread.Sleep(2000);
                try
                {
                    string pdf1 = @"C:\Reportes\ReportePdf_" + "Errores" + "_" + Hora();
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Informe").Coordinates);
                    desktopSession.Mouse.Click(null);
                    newFunctions_4.ScreenshotNuevo("Informe del Proceso", true, file);
                    Thread.Sleep(500);
                    newFunctions_5.GenerarReportes("Errores", file, pdf1);
                }
                catch
                {

                }
            }
        }
    }
}
