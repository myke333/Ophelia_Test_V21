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
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBi
{
    class CrudKbiedfor
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

        public static void CrudKBiedfor(WindowsDriver<WindowsElement> rootSession, int Tipo, List<string> Lupas, string fechaIni, string fechaFin, string editCabeza, string motor)
        {
            if (Tipo == 1)
            {
                if(motor == "SQL")
                {
                    PruebaCRUD.LupaDinamica(rootSession, Lupas);
                    Thread.Sleep(1000);
                }
                else
                {
                    var ElementList = rootSession.FindElementsByClassName("TDBEdit");
                    ElementList[5].SendKeys(Lupas[0]);
                    ElementList[12].SendKeys(Lupas[1]);
                    ElementList[10].SendKeys(Lupas[2]);
                }
                rootSession.Mouse.MouseMove(rootSession.FindElementByName("Datos Adicionales").Coordinates);
                rootSession.Mouse.Click(null);

                var ElementList1 = rootSession.FindElementsByClassName("TDBEdit");

                ElementList1[8].SendKeys(fechaIni);
                Thread.Sleep(500);
                ElementList1[7].SendKeys(fechaFin);
                Thread.Sleep(500);
            }
            else if (Tipo == 2)
            {
                var ElementList = rootSession.FindElementsByClassName("TDBEdit");

                ElementList[7].Clear();
                ElementList[7].SendKeys(editCabeza);
            }
        }

        public static void CrudKBiedforDetalle(WindowsDriver<WindowsElement> rootSession, int Tipo, string tipo, string req, string espec, string editar)
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
    }
}
