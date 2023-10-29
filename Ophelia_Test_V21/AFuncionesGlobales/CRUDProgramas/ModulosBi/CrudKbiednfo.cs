using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBi
{
    class CrudKbiednfo
    {
        private static WindowsDriver<WindowsElement> RootSession()
        {
            var appCapabilities = new AppiumOptions();
            appCapabilities.AddAdditionalCapability("app", "Root");
            var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
            return ApplicationSession;
        }

        static private WindowsDriver<WindowsElement> ReloadSession(WindowsDriver<WindowsElement> session, string Clase)
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

        public static void CrudKBiednfo(WindowsDriver<WindowsElement> rootSession, int Tipo, List<string> Lupas, string nomInsti, string nomEspe, string fechaIni, string editCabeza)
        {
            if (Tipo == 1)
            {
                PruebaCRUD.LupaDinamica(rootSession, Lupas);
                Thread.Sleep(1000);

                var ElementList = rootSession.FindElementsByClassName("TDBEdit");
                var ElementList1 = rootSession.FindElementsByClassName("TDBMemo");
                var check = rootSession.FindElementByName("Externo");
                ElementList[10].SendKeys(nomInsti);
                Thread.Sleep(500);
                ElementList[13].SendKeys(fechaIni);
                Thread.Sleep(1000);
                rootSession.Mouse.MouseMove(check.Coordinates);
                rootSession.Mouse.Click(null);
                Thread.Sleep(500);
                ElementList1[0].SendKeys(nomEspe);
            }
            else if (Tipo == 2)
            {
                var ElementList = rootSession.FindElementsByClassName("TDBMemo");

                ElementList[0].Clear();
                ElementList[0].SendKeys(editCabeza);
            }
        }

        public static void CrudKBiednfoDetalle(WindowsDriver<WindowsElement> rootSession, int Tipo, string tipo, string req, string espec, string editar)
        {
            var ElementList = rootSession.FindElementsByClassName("TDBGrid");

            if (Tipo == 1)
            {
                rootSession.Mouse.MouseMove(ElementList[0].Coordinates, 33, 29);
                rootSession.Mouse.DoubleClick(null);
                //rootSession.Mouse.Click(null);

                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TCaptura");
                var elements = rootSession.FindElementsByClassName("TEdit");

                elements[1].SendKeys(tipo);
                Thread.Sleep(1000);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);

                rootSession = RootSession();
                ElementList = rootSession.FindElementsByClassName("TDBGrid");
                rootSession.Mouse.MouseMove(ElementList[0].Coordinates, 33, 29);
                rootSession.Mouse.Click(null);
                rootSession.Mouse.Click(null);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                rootSession.Keyboard.SendKeys(req);
                Thread.Sleep(500);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                rootSession.Keyboard.SendKeys(espec);
            }
            else if (Tipo == 2)
            {
                rootSession.Mouse.MouseMove(ElementList[0].Coordinates, 33, 29);
                rootSession.Mouse.Click(null);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                rootSession.Keyboard.SendKeys(editar);
            }
        }
    }
}
