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
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_4
{
    class CrudKnmausen : FuncionesVitales
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

                return ApplicationSession;
            }
            catch (Exception ex) { return null; }
        }

        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string file)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList1 = desktopSession.FindElementsByClassName("TBitBtn");
            Debug.WriteLine(ElementList1.Count);
            if (tipo == 1)
            {
                LupaDinamicaNew(desktopSession, data);
                ElementList[24].SendKeys(data[3]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                ElementList1[2].Click();
                WindowsDriver<WindowsElement> rootSession = null;
                rootSession = RootSession();
                rootSession = PruebaCRUD.ReloadSession(rootSession, "TfrmCambioFechas");
                rootSession.Keyboard.SendKeys(data[4]);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                rootSession.Keyboard.SendKeys(data[5]);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                var ElementList3 = desktopSession.FindElementsByClassName("TGroupButton");
                foreach (var elem in ElementList3)
                {
                    if (data[6] == "N" && elem.Text == "No Remunerado")
                    {
                        elem.Click();
                    }
                }
            }
            else
            {
                var ElementList4 = desktopSession.FindElementsByClassName("TGroupButton");
                foreach (var elem1 in ElementList4)
                {
                    if (data[7] == "R" && elem1.Text == "Remunerado")
                    {
                        elem1.Click();
                    }
                }
            }
        }

        public static void LupaDinamicaNew(WindowsDriver<WindowsElement> desktopSession, List<string> data)
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
                if (indexVal != 0)
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
                    }
                }
                indexVal++;
            }
        }

        public static List<string> ReportePreliminar(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1, int tipo, string nomPrograma = null)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            if (bandera == "0" || bandera == "1")
            {
                var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
                {
                    Timeout = TimeSpan.FromSeconds(60),
                    PollingInterval = TimeSpan.FromSeconds(1)
                };
                wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
                WindowsDriver<WindowsElement> rootSession = null;
                WindowsElement codPanel = null;
                bool brk = false;
                bool IsDisplayedPreWin = false;
                wait.Until(driver =>
                {
                    List<Point> coordenadas = FuncionesGlobales.coordinatesFinder(desktopSession, 227, 117, 0);
                    int coordX = coordenadas[0].X;
                    int coordY = coordenadas[0].Y;
                    var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
                    Thread.Sleep(1000);

                    try
                    {
                        desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordX, coordY);
                        desktopSession.Mouse.Click(null);
                        Thread.Sleep(2000);
                        if (bandera == "0")
                        {
                            if (tipo == 1)
                            {
                                WindowsDriver<WindowsElement> rootSession1 = null;
                                rootSession1 = RootSession();
                                newFunctions_4.ScreenshotNuevo("Preliminar Normal", true, file);
                                rootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1, nomPrograma);
                                if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                            }
                            else
                            {
                                WindowsDriver<WindowsElement> rootSession1 = null;
                                rootSession1 = RootSession();
                                rootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                                newFunctions_4.ScreenshotNuevo("Preliminar Totalizado", true, file);
                                rootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1, nomPrograma);
                                if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                            }
                            
                        }
                    }
                    catch (Exception ex)
                    {
                        errorMessages.Add($"Error de Appium: {ex.ToString()}");
                    }
                    codPanel = parentElement;
                    return codPanel != null;
                });
            }
            return errorMessages;
        }

    }
}
