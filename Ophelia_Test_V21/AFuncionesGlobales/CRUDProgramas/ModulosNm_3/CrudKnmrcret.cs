using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_3
{
    class CrudKnmrcret : FuncionesVitales
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
        public static void CrudKNmrcret(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (Tipo == 1)
            {
                PruebaCRUD.LupaDinamica2(desktopSession, data);
                Thread.Sleep(500);
                ElementList[11].SendKeys(data[1]);
                Thread.Sleep(500);
                desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Variable").Coordinates);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                ElementList[8].SendKeys(data[2]);
            }
            else if (Tipo == 2)
            {
                ElementList[8].Clear();
                Thread.Sleep(500);
                ElementList[8].SendKeys(data[3]);
            }
        }
        public static List<string> PreliminarKNmrcret(int Tipo, string pdf, string excel, string file)
        {
            List<string> errors = new List<string>();
            var rootSession = RootSession();
            rootSession = ReloadSession(rootSession, "TFrmNmTiprp");
            if (Tipo == 1)
            {
                rootSession.Mouse.MouseMove(rootSession.FindElementByName("Aceptar").Coordinates);
                rootSession.Mouse.Click(null);
                Thread.Sleep(1000);
                errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf);
                return errors;
            }
            else if (Tipo == 2)
            {
                rootSession.Mouse.MouseMove(rootSession.FindElementByName("Archivo Individual Excel (*.xls)").Coordinates);
                rootSession.Mouse.Click(null);
                Thread.Sleep(500);
                rootSession.Mouse.MouseMove(rootSession.FindElementByName("Aceptar").Coordinates);
                rootSession.Mouse.Click(null);
                Thread.Sleep(2000);
                rootSession = RootSession();
                rootSession.Keyboard.SendKeys(excel);
                Thread.Sleep(500);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(10000);
                newFunctions_4.ScreenshotNuevo("Proceso Terminado", true, file);
                rootSession = RootSession();
                Thread.Sleep(500);
                rootSession = ReloadSession(rootSession, "TFrmInfoP50");
                rootSession.Mouse.MouseMove(rootSession.FindElementByName("Aceptar").Coordinates);
                rootSession.Mouse.Click(null);
                Thread.Sleep(2000);
                try
                {
                    Process.Start(excel + ".xls");
                    Thread.Sleep(4000);
                    rootSession = RootSession();
                    rootSession = ReloadSession(rootSession, "XLMAIN");
                    newFunctions_4.ScreenshotNuevo("Reporte Preliminar Exportado a Excel", true, file);
                    LimpiarProcesos();
                }
                catch
                {
                    errors.Add("El Excel no se encontró con el nombre indicado al momento de guardarlo");
                }
                return errors;
            }
            else if (Tipo == 3)
            {
                rootSession.Mouse.MouseMove(rootSession.FindElementByName("Archivo Individual PDF (*.pdf)").Coordinates);
                rootSession.Mouse.Click(null);
                Thread.Sleep(500);
                rootSession.Mouse.MouseMove(rootSession.FindElementByName("Aceptar").Coordinates);
                rootSession.Mouse.Click(null);
                Thread.Sleep(2000);
                rootSession = RootSession();
                rootSession.Keyboard.SendKeys(pdf);
                Thread.Sleep(500);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(5000);
                newFunctions_4.ScreenshotNuevo("Proceso terminado", true, file);
                rootSession = RootSession();
                Thread.Sleep(500);
                rootSession = ReloadSession(rootSession, "TFrmInfoP50");
                rootSession.Mouse.MouseMove(rootSession.FindElementByName("Aceptar").Coordinates);
                rootSession.Mouse.Click(null);
                Thread.Sleep(2000);
                try
                {
                    Process.Start(pdf + ".pdf");
                    Thread.Sleep(4000);
                    rootSession = RootSession();
                    rootSession = ReloadSession(rootSession, "AcrobatSDIWindow");
                    newFunctions_4.ScreenshotNuevo("Reporte Preliminar Exportado a PDF", true, file);
                    LimpiarProcesos();
                }
                catch
                {
                    errors.Add("El PDF no se encontró con el nombre indicado al momento de guardarlo");
                }
                return errors;
            }
            return errors;
        }
    }
}
