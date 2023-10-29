using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBi
{
    class CrudKbiemple
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

        public static void CrudKBiemple(WindowsDriver<WindowsElement> rootSession, int Tipo, string Id, string PrimerA, string SegundoA, string PrimerN, string SegundoN, string Pais, string Departamento, string Municipio, string FechaNacimiento, string Telefono, string num1, string num2, string num3, string correo, string titulo, string editar)
        {
            var ElementList = rootSession.FindElementsByClassName("TDBEdit");
            var ElementList1 = rootSession.FindElementsByClassName("TEdit");

            if (Tipo == 1)
            {
                ElementList[2].SendKeys(Id);
                Thread.Sleep(500);
                ElementList1[3].SendKeys(PrimerA);
                Thread.Sleep(500);
                ElementList1[1].SendKeys(SegundoA);
                Thread.Sleep(500);
                ElementList1[2].SendKeys(PrimerN);
                Thread.Sleep(500);
                ElementList1[0].SendKeys(SegundoN);
                Thread.Sleep(500);
                ElementList[13].SendKeys(Pais);
                Thread.Sleep(500);
                ElementList[12].SendKeys(Departamento);
                Thread.Sleep(500);
                ElementList[11].SendKeys(Municipio);
                Thread.Sleep(500);
                ElementList[21].SendKeys(Pais);
                Thread.Sleep(500);
                ElementList[20].SendKeys(Departamento);
                Thread.Sleep(500);
                ElementList[19].SendKeys(Municipio);
                Thread.Sleep(500);
                Debugger.Launch();
                ElementList[24].SendKeys(Telefono);
                Thread.Sleep(2000);
                rootSession.Mouse.MouseMove(ElementList[23].Coordinates);
                rootSession.Mouse.Click(null);
                Thread.Sleep(2000);
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TFrmDireccion");

                var Direccion1 = rootSession.FindElementsByName("Abrir");
                var Direccion2 = rootSession.FindElementsByClassName("TEdit");

                rootSession.Mouse.MouseMove(Direccion1[7].Coordinates);
                rootSession.Mouse.Click(null);
                Thread.Sleep(1000);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);

                rootSession.Mouse.MouseMove(Direccion2[4].Coordinates);
                rootSession.Mouse.Click(null);
                Direccion2[4].SendKeys(num1);

                rootSession.Mouse.MouseMove(Direccion1[6].Coordinates);
                rootSession.Mouse.Click(null);
                Thread.Sleep(1000);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);

                rootSession.Mouse.MouseMove(Direccion2[3].Coordinates);
                rootSession.Mouse.Click(null);
                Direccion2[3].SendKeys(num2);
                Thread.Sleep(500);
                Direccion2[2].SendKeys(num3);

                rootSession.Mouse.MouseMove(rootSession.FindElementByName("Aceptar").Coordinates);
                rootSession.Mouse.Click(null);
                Thread.Sleep(1000);

                rootSession = RootSession();
                rootSession.Mouse.MouseMove(rootSession.FindElementByName("Soltero").Coordinates);
                rootSession.Mouse.Click(null);
                Thread.Sleep(1000);

                rootSession.Mouse.MouseMove(rootSession.FindElementByName("Correo Electrónico").Coordinates);
                rootSession.Mouse.Click(null);
                Thread.Sleep(1000);

                rootSession = RootSession();
                var ListaCorreo = rootSession.FindElementsByClassName("TDBEdit");
                ListaCorreo[15].SendKeys(correo);

                rootSession.Mouse.MouseMove(rootSession.FindElementByName("Formación Académica").Coordinates);
                rootSession.Mouse.Click(null);
                Thread.Sleep(1000);

                rootSession = RootSession();
                var ListaTitulo = rootSession.FindElementsByClassName("TDBEdit");
                ListaTitulo[8].SendKeys(titulo);
                Thread.Sleep(1000);

                rootSession.Mouse.MouseMove(rootSession.FindElementByName("Datos Básicos ").Coordinates);
                rootSession.Mouse.Click(null);
                Thread.Sleep(1000);

                rootSession = RootSession();
                var ElementList2 = rootSession.FindElementsByClassName("TDBEdit");

                ElementList2[21].SendKeys(Pais);
                Thread.Sleep(500);
                ElementList2[20].SendKeys(Departamento);
                Thread.Sleep(500);
                ElementList2[19].SendKeys(Municipio);
                Thread.Sleep(500);
                ElementList2[10].SendKeys(FechaNacimiento);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);

                rootSession.Mouse.MouseMove(rootSession.FindElementByName("Colombia            ").Coordinates);
                rootSession.Mouse.Click(null);

                rootSession.Mouse.MouseMove(rootSession.FindElementByName("Masculino").Coordinates);
                rootSession.Mouse.Click(null);
            }
            else if (Tipo == 2)
            {
                ElementList1[1].Clear();
                ElementList1[1].SendKeys(editar);
            }
        }
    }
}
