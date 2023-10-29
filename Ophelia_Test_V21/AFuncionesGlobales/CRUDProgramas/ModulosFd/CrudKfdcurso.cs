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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosFd
{
    class CrudKfdcurso:FuncionesVitales
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
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (tipo == 1)
            {
                ElementList[6].SendKeys(data[0]);
                ElementList[5].SendKeys(data[1]);
                LupaDinamica(desktopSession, data);
                ElementList[2].SendKeys(data[3]);
                newFunctions_4.ScreenshotNuevo("Agregar Registro", true, file);
            }
            else
            {
                ElementList[5].Clear();
                ElementList[5].SendKeys(data[4]);
                newFunctions_4.ScreenshotNuevo("Editar Registro", true, file);
            }

        }
        public static void CrudDetalle(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string file)
        {
            WindowsDriver<WindowsElement> rootSession = null;
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            if (tipo == 1)
            {

                Thread.Sleep(2000);
                for (int i = 0; i < 3; i++)
                {
                    if (i == 0)
                    {
                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 34, 31);
                        desktopSession.Mouse.Click(null);
                        Thread.Sleep(500);
                    }
                    else if (i == 1)
                    {
                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 434, 31);
                        desktopSession.Mouse.Click(null);
                        Thread.Sleep(500);
                    }
                    else
                    {
                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 795, 31);
                        desktopSession.Mouse.Click(null);
                        Thread.Sleep(500);
                    }
                    desktopSession.Mouse.Click(null);
                    desktopSession.Mouse.DoubleClick(null);
                    rootSession = RootSession();
                    rootSession = ReloadSession(rootSession, "TCaptura");
                    Thread.Sleep(1000);
                    var campo = rootSession.FindElementsByClassName("TEdit");
                    ////Debugger.Launch();
                    rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                    rootSession.Mouse.Click(null);
                    campo[1].Clear();
                    rootSession.Keyboard.SendKeys(data[i]);
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
                   
                }
                newFunctions_4.ScreenshotNuevo("Agregar Detalle", true, file);
            }
            else
            {
                ElementList[5].Clear();
                ElementList[5].SendKeys(data[4]);
            }

        }
        public static void EliminarDetalle(WindowsDriver<WindowsElement> desktopSession, WindowsElement ElementList, Dictionary<string, Point> botonesNavegador, string file)
        {
            WindowsDriver<WindowsElement> rootSession = null;
            desktopSession.Mouse.MouseMove(ElementList.Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Borrar Detalle", true, file);
            rootSession = RootSession();
            rootSession = ReloadSession(rootSession, "TMessageForm");
            rootSession.Mouse.MouseMove(rootSession.FindElementByName("OK").Coordinates);
            rootSession.Mouse.Click(null);
            newFunctions_4.ScreenshotNuevo("Detalle 1 Eliminado", true, file);
            Thread.Sleep(2000);
        }
        public static void LupaDinamica(WindowsDriver<WindowsElement> desktopSession, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            List<Point> coordenadas = PruebaCRUD.CoordinatesFinder(desktopSession, 82, 150, 75);
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            int contLupa = 1;
            int indexVal = 2;
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
