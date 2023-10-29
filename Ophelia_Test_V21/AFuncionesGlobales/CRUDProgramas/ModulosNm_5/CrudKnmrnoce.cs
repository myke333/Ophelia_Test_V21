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
namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_5
{
    class CrudKnmrnoce : FuncionesVitales
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

        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TEdit");
            ElementList[1].SendKeys(data[0]);
            ElementList[0].SendKeys(data[1]);
            ElementList[3].SendKeys(data[2]);
            ElementList[4].SendKeys(data[3]);
        }


        public static void ButtonGenerarDatos(WindowsDriver<WindowsElement> desktopSession, string ruta, string file)
        {
            List<string> errorMessages = new List<string>();
            desktopSession.FindElementByName("Generar Datos").Click();
            WindowsDriver<WindowsElement> Session = null;
            Session = RootSession();
            Session.Keyboard.SendKeys(ruta);
            newFunctions_4.ScreenshotNuevo("Reporte Generar Datos Guardar", true, file);
            Session.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(1000);
            

        }

        public static void Ventana(WindowsDriver<WindowsElement> desktopSession, string file)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            newFunctions_4.ScreenshotNuevo("Proceso Finalizado", true, file);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
        }

        public static void VentanaWord(WindowsDriver<WindowsElement> desktopSession)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
        }

        public static void VentanaWord2(WindowsDriver<WindowsElement> desktopSession, string file)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            newFunctions_4.ScreenshotNuevo("Ver Documento", true, file);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

            rootSession = RootSession();
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            newFunctions_4.ScreenshotNuevo("Cesantias", true, file);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(5000);
            newFunctions_4.ScreenshotNuevo("Anuncio Word", true, file);
            rootSession = RootSession();
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowLeft);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(3000);
            newFunctions_4.ScreenshotNuevo("Letrero No Encuentra Origen de Datos Word", true, file);
            LimpiarProcesos();

        }

        public static void VentanaWord3(WindowsDriver<WindowsElement> desktopSession, string file)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            newFunctions_4.ScreenshotNuevo("Documento Word", true, file);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

            rootSession = RootSession();
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            newFunctions_4.ScreenshotNuevo("Intereses", true, file);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2000);

            newFunctions_4.ScreenshotNuevo("Error al recuperar el archivo", true, file);
            rootSession = RootSession();
            PruebaCRUD.VentanaAzul(rootSession);

        }

    }
}
