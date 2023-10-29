using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSo
{
    class CrudKsotipin
    {
        public static void CrudKSotipin(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (Tipo == 1)
            {
                ElementList[3].SendKeys(data[0]);
                Thread.Sleep(500);
                ElementList[2].SendKeys(data[1]);
                Thread.Sleep(500);
            }
            else if (Tipo == 2)
            {
                ElementList[2].Clear();
                Thread.Sleep(500);
                ElementList[2].SendKeys(data[2]);
                Thread.Sleep(500);
            }
        }
        public static void CrudKSotipinDetalle1(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, 66, 26);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(500);
            if (Tipo == 1)
            {
                ElementList[1].SendKeys(data[0]);
                Thread.Sleep(200);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(200);
                ElementList[1].SendKeys(data[1]);
                Thread.Sleep(500);
            }
            else if (Tipo == 2)
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(200);
                ElementList[1].Clear();
                Thread.Sleep(500);
                ElementList[1].SendKeys(data[2]);
                Thread.Sleep(500);
            }
        }
        public static void CrudKSotipinDetalle2(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 66, 26);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(500);
            if (Tipo == 1)
            {
                ElementList[0].SendKeys(data[0]);
                Thread.Sleep(200);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(200);
                ElementList[0].SendKeys(data[1]);
                Thread.Sleep(500);
            }
            else if (Tipo == 2)
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(200);
                ElementList[0].Clear();
                Thread.Sleep(500);
                ElementList[0].SendKeys(data[2]);
                Thread.Sleep(500);
            }
        }
        public static void CrudKSotipinDetalle3(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 66, 26);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(500);
            if (Tipo == 1)
            {
                ElementList[0].SendKeys(data[0]);
                Thread.Sleep(200);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(200);
                ElementList[0].SendKeys(data[1]);
                Thread.Sleep(200);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(200);
                ElementList[0].SendKeys(data[2]);
                Thread.Sleep(500);
            }
            else if (Tipo == 2)
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(200);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(200);
                ElementList[0].Clear();
                Thread.Sleep(500);
                ElementList[0].SendKeys(data[3]);
                Thread.Sleep(500);
            }
        }
        public static List<string> CrudKSotipinPreliminar(WindowsDriver<WindowsElement> desktopSession, string BanderaPreli, string ruta, string file)
        {

            List<string> errores = new List<string> { };
            List<string> errors = new List<string> { };
            WindowsDriver<WindowsElement> rootSession;
            errors = FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, ruta);
            if (errors.Count > 0) { foreach (string er in errors) { errores.Add(er); } }
            Thread.Sleep(1000);
            rootSession = RootSession();
            rootSession = ReloadSession(rootSession, "TFrmOpcRepo");
            rootSession.Mouse.MouseMove(rootSession.FindElementByName("OK").Coordinates);
            rootSession.Mouse.Click(null);
            Thread.Sleep(2000);
            errores = newFunctions_5.GenerarReportes("Preliminar", file, ruta);
            if (errors.Count > 0) { foreach (string er in errors) { errores.Add(er); } }
            Thread.Sleep(1000);
            errors = FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, ruta);
            if (errors.Count > 0) { foreach (string er in errors) { errores.Add(er); } }
            rootSession.Mouse.MouseMove(rootSession.FindElementByName("Red de Apoyo").Coordinates);
            rootSession.Mouse.Click(null);
            Thread.Sleep(500);
            rootSession.Mouse.MouseMove(rootSession.FindElementByName("OK").Coordinates);
            rootSession.Mouse.Click(null);
            Thread.Sleep(2000);
            errores = newFunctions_5.GenerarReportes("Preliminar", file, ruta);
            if (errors.Count > 0) { foreach (string er in errors) { errores.Add(er); } }

            return errores;
        }
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
    }
}
