using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosGn
{
    class Crudkgnrperm:FuncionesVitales
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

        public static List<string> CRUD(WindowsDriver<WindowsElement> desktopSession,string bandera, string file, string pdf1,string ruta,string motor)
        {
            List<string> errors = new List<string>();
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;

            //REPORTE PRELIMINAR
            rootSession = RootSession();
            rootSession = ReloadSession(rootSession, "TFrmMenu");
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            newFunctions_4.ScreenshotNuevo("Opcion Preliminar", true, file);
            Thread.Sleep(500);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            if (motor == "SQL")
            {
                Thread.Sleep(100000);
            }
            else
            {
                Thread.Sleep(500000);
            }
            errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
            if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
            Thread.Sleep(5000);

            //REPORTE EXPORTAR EXCEL
            FuncionesGlobales.ReportePreliminar(desktopSession, bandera, file, pdf1);
            rootSession = RootSession();
            rootSession = ReloadSession(rootSession, "TFrmMenu");
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
            newFunctions_4.ScreenshotNuevo("Opcion Excel", true, file);
            Thread.Sleep(500);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(50000);
            generarExcel(desktopSession, file, ruta);
            return errorMessages;

        }

        public static List<string> generarExcel(WindowsDriver<WindowsElement> Session, string file,/* string codProgram,*/ string ruta)
        {
            List<string> errorMessages = new List<string>();
            try
            {
                Thread.Sleep(5000);
                Session = RootSession();
                Session = ReloadSession(Session, "XLMAIN");
                var saveExcel = Session.FindElementsByName("Maximizar");
                if (saveExcel.Count > 0)
                {
                    saveExcel[1].Click();
                    Session.FindElementByName("Guardar").Click();
                    Session.FindElementByName("Examinar").Click();
                    Session.Keyboard.SendKeys(ruta);
                    newFunctions_4.ScreenshotNuevo("Reporte Excel Guardado", true, file);
                    Session.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);
                    Thread.Sleep(4000);
                    LimpiarProcesos();
                }
            }
            catch
            {
                Session.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                errorMessages.Add($"Error al intentar generar el excel");
            }
            return errorMessages;
        }
    }
}
