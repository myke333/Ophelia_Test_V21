using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosFd
{
    class CrudKfdentid
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
                //appCapabilities.AddAdditionalCapability("app", "Root");
                appCapabilities.AddAdditionalCapability("appTopLevelWindow", ApplicationSessionHandle);
                var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);

                string ses = ApplicationSession.PageSource;
                ////Debugger.Launch();
                return ApplicationSession;
            }
            catch (Exception ex) { return null; }
        }
        public static void CrudKFdentid(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (Tipo == 1)
            {
                ElementList[4].SendKeys(data[0]);
                ElementList[3].SendKeys(data[1]);
                ElementList[2].SendKeys(data[2]);
                ElementList[17].SendKeys(data[3]);
                desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Cédula Ciudadanìa").Coordinates);
                desktopSession.Mouse.Click(null);
            }
            else if (Tipo == 2)
            {
                ElementList[3].Clear();
                ElementList[3].SendKeys(data[4]);
            }
        }
        public static List<string> Preliminar(string file, string ruta)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();

            WindowsDriver<WindowsElement> rootSession;
            rootSession = RootSession();
            rootSession = ReloadSession(rootSession, "TFrmFdRepor");
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Opción Preliminar", true, file);
            Thread.Sleep(500);
            rootSession.Mouse.MouseMove(rootSession.FindElementByName("Aceptar").Coordinates);
            rootSession.Mouse.Click(null);
            Thread.Sleep(2000);
            errors = newFunctions_5.GenerarReportes("Preliminar", file, ruta);
            if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }

            return errorMessages;
        }
    }
}
