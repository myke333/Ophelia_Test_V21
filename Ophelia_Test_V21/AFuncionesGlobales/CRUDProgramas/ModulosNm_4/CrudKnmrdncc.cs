using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Support.UI;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_4
{
    class CrudKnmrdncc : FuncionesVitales
    {
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
        protected static WindowsDriver<WindowsElement> rootSession;
        public static List<string> PreliminarKNmrdncc(WindowsDriver<WindowsElement> desktopSession, string BanderaPreli, string file, string pdf1, string TipoQbe, string QbeFilter)
        {
            List<string> errors = new List<string> { };
            List<string> ErrorValidacion = new List<string> { };
            for (int i = 0; i < 2; i++)
            {
                //QBE
                errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                Thread.Sleep(500);
                FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
                Thread.Sleep(2000);
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TFrmMenu");
                if (i == 1)
                {
                    rootSession.Mouse.MouseMove(rootSession.FindElementByName("Totalizado por Centro de Costo").Coordinates);
                    rootSession.Mouse.Click(null);
                    Thread.Sleep(500);
                }
                rootSession.Mouse.MouseMove(rootSession.FindElementByName("Aceptar").Coordinates);
                rootSession.Mouse.Click(null);
                Thread.Sleep(500);
                try
                {
                    errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                    if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                }
                catch
                {
                    PruebaCRUD.cerrarBorrar(desktopSession);
                    newFunctions_4.ScreenshotNuevo("No abre el preliminar numero: " + (i+1), true, file);
                }
            }
            return ErrorValidacion;
        }
        public static List<string> ImprimirKNmrdncc(WindowsDriver<WindowsElement> desktopSession, string file, string TipoQbe, string QbeFilter)
        {
            List<string> errors = new List<string> { };
            List<string> ErrorValidacion = new List<string> { };
            for (int i = 0; i < 2; i++)
            {
                //QBE
                errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                Thread.Sleep(500);
                var printButton = desktopSession.FindElementByName("Imprimir");
                desktopSession.Mouse.MouseMove(printButton.Coordinates, printButton.Size.Width / 2, printButton.Size.Height / 2);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TFrmMenu");
                if (i == 1)
                {
                    rootSession.Mouse.MouseMove(rootSession.FindElementByName("Detallado por Empleado").Coordinates);
                    rootSession.Mouse.Click(null);
                    Thread.Sleep(500);
                }
                rootSession.Mouse.MouseMove(rootSession.FindElementByName("Aceptar").Coordinates);
                rootSession.Mouse.Click(null);
                Thread.Sleep(500);
                try
                {
                    rootSession = RootSession();
                    rootSession = ReloadSession(rootSession, "TfrxPrintDialog");

                    var scrollelement = rootSession.FindElementsByClassName("TComboBox");
                    rootSession.Mouse.MouseMove(scrollelement[5].Coordinates, 10, 10);
                    rootSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Ventana Impresión PDF", true, file);
                    rootSession.Keyboard.SendKeys("m");
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(500);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);
                    string pdf = @"C:\Reportes\ReportePDF1_" + "Imprimir" + "_" + Hora();
                    rootSession = RootSession();
                    rootSession.Keyboard.SendKeys(pdf);
                    Thread.Sleep(500);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(5000);
                    Process.Start(pdf + ".pdf");
                    Thread.Sleep(6000);
                    newFunctions_4.ScreenshotNuevo("Impresión PDF", true, file);
                    LimpiarProcesos();
                }
                catch
                {
                    PruebaCRUD.cerrarBorrar(desktopSession);
                    newFunctions_4.ScreenshotNuevo("No abre el Imprimir numero: " + (i + 1), true, file);
                }
            }
            return ErrorValidacion;
        }
    }
}
