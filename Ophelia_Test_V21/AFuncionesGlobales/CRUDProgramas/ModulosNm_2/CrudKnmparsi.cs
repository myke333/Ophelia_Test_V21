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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_2
{
    class CrudKnmparsi:FuncionesVitales
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

        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data,  string file)
        {
            if(tipo==1)
            {
                LupaDinamica(desktopSession, data, tipo);
                Thread.Sleep(500);
                newFunctions_4.ScreenshotNuevo("Agregar Registro", true, file);
            }
            else
            {
                LupaDinamica(desktopSession, data, tipo);
                newFunctions_4.ScreenshotNuevo("Editar Registro", true, file);
            }
            
        }

        public static void LupaDinamica(WindowsDriver<WindowsElement> desktopSession, List<string> data, int tipo)
        {
            List<string> errorMessages = new List<string>();
            List<Point> coordenadas = PruebaCRUD.CoordinatesFinder(desktopSession, 82, 150, 75);
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            int contLupa = 1;
            bool IsDisplayedQbe = false;
            int indexVal = 0;
            foreach (Point coord in coordenadas)
            {
                Thread.Sleep(500);
                if (tipo == 1)
                {
                    if (contLupa == 1 || contLupa == 11 || contLupa == 14)
                    {
                        desktopSession.Mouse.MouseMove(parentElement.Coordinates, coord.X + 10, coord.Y);
                        desktopSession.Mouse.Click(null);


                        if (contLupa == 14)
                        {
                            try
                            {
                                rootSession = RootSession();
                                rootSession = PruebaCRUD.ReloadSession(rootSession, "TKQBE");
                            }
                            catch
                            {

                            }
                            if (rootSession != null)
                            {
                                IsDisplayedQbe = true;
                            }
                            else
                            {
                                errorMessages.Add($"No puede encontrar la ventana de QBE");
                            }
                            if (IsDisplayedQbe)
                            {
                                var elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                                if (!string.IsNullOrEmpty(data[indexVal]))
                                {
                                    elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/Pane[contains(@ClassName, 'TStringGrid')]/*"));
                                    elements[0].SendKeys(data[indexVal]);
                                }
                                elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                                new Actions(rootSession).MoveToElement(elements[8], 0, 0).MoveByOffset(20, 10).DoubleClick().Perform();
                                Thread.Sleep(500);
                                IsDisplayedQbe = false;

                            }
                        }
                        rootSession = RootSession();
                        rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");
                        //var campo = rootSession.FindElementsByXPath("//Window[@ClassName=\"TCaptura\"][@Name=\" LISTA DE EMPLEADOS\"]/Edit[@ClassName=\"TEdit\"]");
                        var campo = rootSession.FindElementsByClassName("TEdit");
                        ////Debugger.Launch();
                        rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                        rootSession.Mouse.Click(null);
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
                    contLupa++;
                }
                else
                {
                    if (contLupa==1)
                    {
                        desktopSession.Mouse.MouseMove(parentElement.Coordinates, coord.X + 10, coord.Y);
                        desktopSession.Mouse.Click(null);

                        rootSession = RootSession();
                        rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");
                        //var campo = rootSession.FindElementsByXPath("//Window[@ClassName=\"TCaptura\"][@Name=\" LISTA DE EMPLEADOS\"]/Edit[@ClassName=\"TEdit\"]");
                        var campo = rootSession.FindElementsByClassName("TEdit");
                        ////Debugger.Launch();
                        rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                        rootSession.Mouse.Click(null);
                        rootSession.Keyboard.SendKeys(data[data.Count-1]);
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
                        contLupa++;
                    }
                }
            }

        }








    }
}
