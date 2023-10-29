using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBi
{
    class CrudKbihvext
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
                //appCapabilities.AddAdditionalCapability("app", "Root");
                appCapabilities.AddAdditionalCapability("appTopLevelWindow", ApplicationSessionHandle);
                var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);

                string ses = ApplicationSession.PageSource;
                ////Debugger.Launch();
                return ApplicationSession;
            }
            catch (Exception ex) { return null; }
        }

        public static void CrudKBihvext(WindowsDriver<WindowsElement> rootSession, int Tipo, List<string> Lupas, string nomEmpr, string fechaIng, string pais, string dep, string mun, string editCabeza)
        {
            if (Tipo == 1)
            {
                PruebaCRUD.LupaDinamica(rootSession, Lupas);
                Thread.Sleep(1000);
                
                var ElementList = rootSession.FindElementsByClassName("TDBEdit");
                ElementList[11].SendKeys(nomEmpr);
                Thread.Sleep(500);
                ElementList[28].SendKeys(fechaIng);
                Thread.Sleep(500);
                ElementList[26].SendKeys(pais);
                Thread.Sleep(500);
                ElementList[25].SendKeys(dep);
                Thread.Sleep(500);
                ElementList[24].SendKeys(mun);

                rootSession.Mouse.MouseMove(rootSession.FindElementByName("Empleo Actual").Coordinates);
                rootSession.Mouse.Click(null);
                Thread.Sleep(1000);
                rootSession.Mouse.MouseMove(rootSession.FindElementByName("Privada").Coordinates);
                rootSession.Mouse.Click(null);

            }
            else if (Tipo == 2)
            {
                var ElementList = rootSession.FindElementsByClassName("TDBEdit");
                ElementList[11].Clear();
                ElementList[11].SendKeys(editCabeza);
                Thread.Sleep(500);
            }
        }

        public static void CrudKBihvextDetalle(WindowsDriver<WindowsElement> rootSession, int Tipo, string tipo, string req, string espec, string editar)
        {
            var ElementList = rootSession.FindElementsByClassName("TDBGrid");

            if (Tipo == 1)
            {
                rootSession.Mouse.MouseMove(ElementList[0].Coordinates, 33, 29);
                rootSession.Mouse.Click(null);
                rootSession.Mouse.Click(null);

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

        public static void CrudKBihvextDetalle1(WindowsDriver<WindowsElement> rootSession, int Tipo, string codigo, string editDetalle)
        {
            var ElementList = rootSession.FindElementsByClassName("TDBGrid");

            if (Tipo == 1)
            {
                rootSession.Mouse.MouseMove(ElementList[0].Coordinates, 33, 29);
                rootSession.Mouse.Click(null);

                ElementList[0].SendKeys(codigo);
                Thread.Sleep(1000);
            }
            else if (Tipo == 2)
            {
                rootSession.Mouse.MouseMove(ElementList[0].Coordinates, 33, 29);
                rootSession.Mouse.Click(null);

                ElementList[0].Clear();
                ElementList[0].SendKeys(editDetalle);
                Thread.Sleep(1000);
            }
        }
        public static void CrudKBihvextDetalle2(WindowsDriver<WindowsElement> rootSession, int Tipo, string herra, string nivel, string cal, string observa, string editDetalle)
        {
            var herramientas = rootSession.FindElementsByName("Herramientas");
            herramientas[1].Click();
            Thread.Sleep(1000);
            var ElementList = rootSession.FindElementsByClassName("TDBGrid");
            
            if (Tipo == 1)
            {
                rootSession.Mouse.MouseMove(ElementList[0].Coordinates, 33, 29);
                rootSession.Mouse.Click(null);

                ElementList[0].SendKeys(herra);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                ElementList[0].SendKeys(nivel);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                ElementList[0].SendKeys(cal);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                ElementList[0].SendKeys(cal);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                ElementList[0].SendKeys(cal);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                ElementList[0].SendKeys(observa);
            }
            else if (Tipo == 2)
            {
                rootSession.Mouse.MouseMove(ElementList[0].Coordinates, 33, 29);
                rootSession.Mouse.Click(null);

                for(int i = 0; i < 5; i++)
                {
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                }

                ElementList[0].Clear();
                ElementList[0].SendKeys(editDetalle);
                Thread.Sleep(1000);
            }
        }
    }
}
