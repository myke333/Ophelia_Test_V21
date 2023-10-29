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
    class CrudKnmrlidd : FuncionesVitales
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
            ElementList[0].SendKeys(data[1]);
            ElementList[1].SendKeys(data[0]);
        }


        public static void Ventana(WindowsDriver<WindowsElement> desktopSession)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
        }

        public static void ButtonImprimir(WindowsDriver<WindowsElement> desktopSession)
        {
            List<string> errorMessages = new List<string>();
            desktopSession.FindElementByName("Imprimir").Click();
        }

        public static void BlocNotas(WindowsDriver<WindowsElement> Session, string file, string ruta)
        {
            Session = RootSession();
            Session = ReloadSession(Session, "Notepad");
            Thread.Sleep(500);
            var saveNotepad = Session.FindElementsByName("Maximizar");
            newFunctions_4.ScreenshotNuevo("Reporte Excel", true, file);
            saveNotepad[0].Click();
            Thread.Sleep(1000);
            Session.FindElementByName("Archivo").Click();
            Session.FindElementByName("Guardar como...").Click(); //KNmRLidd
            Session.Keyboard.SendKeys(ruta);
            Session.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Reporte Preliminar NotePad Guardado", true, file);
            LimpiarProcesos();
        }

    }
}
