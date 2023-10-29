using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosGe
{
    class CrudKgecpeor
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
        public static void CrudKGecpeor(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList1 = desktopSession.FindElementsByClassName("TDBMemo");
            if (Tipo == 1)
            {
                var tipoPlan = desktopSession.FindElementsByName("Abrir");
                desktopSession.Mouse.MouseMove(tipoPlan[1].Coordinates);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("De Meta").Coordinates);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                ElementList[2].SendKeys(data[0]);
                Thread.Sleep(500);
                ElementList1[0].SendKeys(data[1]);
                Thread.Sleep(500);
            }
            else if (Tipo == 2)
            {
                ElementList1[0].Clear();
                Thread.Sleep(500);
                ElementList1[0].SendKeys(data[2]);
                Thread.Sleep(500);
            }
        }
        public static void CrudKGecpeorDetalle(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");

            if (Tipo == 1)
            {
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 33, 29);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(data[0]);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(data[1]);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(data[2]);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(data[3]);
                Thread.Sleep(500);
            }
            else if (Tipo == 2)
            {
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 33, 29);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(data[4]);
                Thread.Sleep(500);
            }
        }
        public static List<string> CrudKGecpeorPreliminar(WindowsDriver<WindowsElement> desktopSession, string BanderaPreli, string file, string pdf1)
        {
            List<string> errorMessages = new List<string>();
            List<string> ErrorValidacion = new List<string>();
            for (int i = 0; i < 2; i++)
            {
                FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
                Thread.Sleep(1000);
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TFrmTipRepo");
                if (i == 0)
                {
                    rootSession.Mouse.MouseMove(rootSession.FindElementByName("Aceptar").Coordinates);
                    rootSession.Mouse.Click(null);
                }
                else
                {
                    rootSession.Mouse.MouseMove(rootSession.FindElementByName("Agrupado por Área").Coordinates);
                    rootSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    rootSession.Mouse.MouseMove(rootSession.FindElementByName("Aceptar").Coordinates);
                    rootSession.Mouse.Click(null);
                }
                
                errorMessages = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                if (errorMessages.Count > 0) { foreach (string er in errorMessages) { ErrorValidacion.Add(er); } }
            }
            
            return ErrorValidacion;

        }
    }
}
