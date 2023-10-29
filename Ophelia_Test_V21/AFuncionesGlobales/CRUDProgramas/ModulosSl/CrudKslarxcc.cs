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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSl
{
    class CrudKslarxcc : FuncionesVitales
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

        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string file)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList1 = desktopSession.FindElementsByClassName("TGroupButton");
            if (tipo == 1)
            {
                LupaDinamica(desktopSession, data);
                foreach (var radiogroup in ElementList1)
                {
                    if (radiogroup.Text == "Integral" || radiogroup.Text == "Indefinido" || radiogroup.Text == "No")
                    {
                        radiogroup.Click();
                    }
                }
            }
            else
            {
                ElementList[5].Clear();
                ElementList[5].SendKeys(data[3]);
            }
        }

        public static void CrudDetalle(WindowsDriver<WindowsElement> desktopSession, int nummDetalle, string file)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            switch (nummDetalle)
            {
                case 1:
                    Thread.Sleep(2000);
                    desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 51, 37);
                    desktopSession.Mouse.DoubleClick(null);
                    obtenerVentana(desktopSession);
                    break;
                case 2:
                    Thread.Sleep(2000);
                    desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 37, 32);
                    desktopSession.Mouse.DoubleClick(null);
                    obtenerVentana(desktopSession);
                    break;
                case 3:
                    Thread.Sleep(2000);
                    desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 32, 35);
                    desktopSession.Mouse.DoubleClick(null);
                    obtenerVentana(desktopSession);
                    break;
                default:
                    Debug.WriteLine("Error la bandera que seleccionas no pertenece a un número de detalle especifico");
                    break;
            }
        }

        public static void obtenerVentana(WindowsDriver<WindowsElement> desktopSession)
        {
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2000);
        }

        public static void ClickButton(WindowsDriver<WindowsElement> desktopSession, int tipo)
        {
            List<string> errorMessages = new List<string>();
            if (tipo == 1)
            {
                desktopSession.FindElementByName("Recursos para el Cargo").Click();
            }
            else if (tipo == 2)
            {
                desktopSession.FindElementByName("Identificación").Click();
            } 
            else if(tipo == 3)
            {
                desktopSession.FindElementByName("Información de Permisos y Accesos").Click();
            }
            else
            {
                desktopSession.FindElementByName("Activos Fijos").Click();
            }
        }

        public static void LupaDinamica(WindowsDriver<WindowsElement> desktopSession, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            List<Point> coordenadas = PruebaCRUD.CoordinatesFinder(desktopSession, 82, 150, 75);
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            int contLupa = 1;
            int indexVal = 0;
            foreach (Point coord in coordenadas)
            {
                desktopSession.Mouse.MouseMove(parentElement.Coordinates, coord.X + 10, coord.Y);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                if (contLupa == 1)
                {
                    rootSession = RootSession();
                    rootSession = ReloadSession(rootSession, "TCaptura");
                    Thread.Sleep(1000);
                    var campo = rootSession.FindElementsByClassName("TEdit");
                    ////Debugger.Launch();
                    rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                    rootSession.Mouse.Click(null);
                    campo[1].Clear();
                    rootSession.Keyboard.SendKeys(data[indexVal]);
                    //Aqui se puede enviar un valor

                    var btn = rootSession.FindElementsByClassName("TBitBtn");
                    ////Debugger.Launch();
                    foreach (var boton in btn)
                    {
                        if (boton.Text == "Aceptar")
                        {
                            rootSession.Mouse.MouseMove(boton.Coordinates, 10, 10);
                            rootSession.Mouse.Click(null);
                        }
                    }
                    indexVal++;
                }
            }
        }

    }
}
