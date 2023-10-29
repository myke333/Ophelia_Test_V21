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
    class CrudKnmrlide : FuncionesVitales
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

        public static void VentanaPreliminar(WindowsDriver<WindowsElement> desktopSession, string file, int tipo)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            if (tipo == 1)
            {
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                newFunctions_4.ScreenshotNuevo("Preliminar Liquidación Definitiva", true, file);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(3000);
                newFunctions_4.ScreenshotNuevo("Proceso Terminado", true, file);
            }
            else
            {
                for (int i = 0; i < 4; i++)
                {
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                }
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                newFunctions_4.ScreenshotNuevo("Preliminar Implementos Devolutivos", true, file);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
        }

        public static void Preview(WindowsDriver<WindowsElement> Session, string file, string ruta)
        {

            Thread.Sleep(5000);
            Session = RootSession();
            Session = ReloadSession(Session, "TQRStandardPreview");
            Thread.Sleep(1000);
            var savePreli = Session.FindElementsByName("Maximizar");
            savePreli[0].Click();
            newFunctions_4.ScreenshotNuevo("Reporte Preview Preliminar", true, file); 
            Thread.Sleep(1000);
            //Guardar
            var ElementList = Session.FindElementsByClassName("TToolBar");
            Debug.WriteLine(ElementList.Count);
            Thread.Sleep(2000);
            Session.Mouse.MouseMove(ElementList[0].Coordinates, 562, 12);
            Session.Mouse.Click(null);
            Thread.Sleep(2000);
            //Ruta
            Session = RootSession();
            Session.Keyboard.SendKeys(ruta);
            newFunctions_4.ScreenshotNuevo("Reporte Preview Preliminar Guardar", true, file);
            Session.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(1000);
            //newFunctions_4.ScreenshotNuevo("Reporte Preview Preliminar Guardado", true, file);
            //Cerrar
            Thread.Sleep(2000);
            Session = RootSession();
            Session = ReloadSession(Session, "TQRStandardPreview");
            var closePreli = Session.FindElementsByName("Cerrar");
            closePreli[0].Click();

        }



        public static void VentanaImprimir(WindowsDriver<WindowsElement> desktopSession, string file, string ruta, int tipo)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            if (tipo == 1)
            {
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowUp);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                newFunctions_4.ScreenshotNuevo("Imprimir Liquidación Definitiva", true, file);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(3000);
                newFunctions_4.ScreenshotNuevo("Proceso Terminado", true, file);

            }
            else
            {
                for (int i = 0; i < 4; i++)
                {
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                }
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                newFunctions_4.ScreenshotNuevo("Imprimir Implementos Devolutivos", true, file);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);
                rootSession = RootSession();
                rootSession.Keyboard.SendKeys(ruta);
                newFunctions_4.ScreenshotNuevo("Guardar Imprimir Implementos Devolutivos", true, file);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

            }
        }

        public static void ClickButtonPreliminar(WindowsDriver<WindowsElement> desktopSession)
        {
            List<string> errorMessages = new List<string>();
            desktopSession.FindElementByName("Preliminar").Click();
        }

        public static void ClickButtonImprimir(WindowsDriver<WindowsElement> desktopSession)
        {
            List<string> errorMessages = new List<string>();
            desktopSession.FindElementByName("Imprimir").Click();
        }

        public static void ClickButtonErroresr(WindowsDriver<WindowsElement> desktopSession)
        {
            List<string> errorMessages = new List<string>();
            desktopSession.FindElementByName("Errores").Click();
        }

        public static void Ventana(WindowsDriver<WindowsElement> desktopSession)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
        }



    }
}
