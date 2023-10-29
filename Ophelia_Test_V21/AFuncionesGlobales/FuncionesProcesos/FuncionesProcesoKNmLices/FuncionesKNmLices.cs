using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using Ophelia_Test_V21.AFuncionesGlobales;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using OpenQA.Selenium;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Support.UI;
using System.Drawing;

namespace Ophelia_Test_V21.AFuncionesGlobales.FuncionesProcesos.FuncionesProcesoKNmLices
{
    class FuncionesKNmLices : FuncionesVitales
    {
        static public WindowsDriver<WindowsElement> RootSession()
        {

            var appCapabilities = new AppiumOptions();
            appCapabilities.AddAdditionalCapability("app", "Root");
            var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
            return ApplicationSession;
        }
        static public WindowsDriver<WindowsElement> ReloadQbeSession(WindowsDriver<WindowsElement> session, string className)
        {
            try
            {
                var ApplicationWindow = session.FindElementByClassName(className);
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

        public static List<string> AgregarCesantiasParciales(WindowsDriver<WindowsElement> desktopSession, List<string> crudVars, string file)
        {
            List<string> errorMessages = new List<string>();
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();

            try
            {
                var TDBEdit = desktopSession.FindElementsByClassName("TDBEdit");
                TDBEdit[8].SendKeys(crudVars[0]);
                TDBEdit[8].SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                switch (crudVars[1])
                {
                    //Tipo de Solicitud
                    case ("Parciales Empresa"):
                        var TipoSol = desktopSession.FindElementsByName(crudVars[1]);
                        TipoSol[0].Click();
                        TipoSol[0].Click();
                        break;
                    case ("Definitivas"):
                        TipoSol = desktopSession.FindElementsByName(crudVars[1]);
                        TipoSol[0].Click();
                        TipoSol[0].Click();
                        break;
                    case ("Otras"):
                        TipoSol = desktopSession.FindElementsByName(crudVars[1]);
                        TipoSol[0].Click();
                        TipoSol[0].Click();
                        break;
                    case ("Parciales Fondo"):
                        TipoSol = desktopSession.FindElementsByName(crudVars[1]);
                        TipoSol[0].Click();
                        TipoSol[0].Click();
                        break;
                }
                //Fecha de Solicitud
                TDBEdit[15].SendKeys(crudVars[2]);
                TDBEdit[15].SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                //Fecha de Corte
                TDBEdit[16].SendKeys(crudVars[3]);
                TDBEdit[16].SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                //Valor solicitado
                TDBEdit[13].SendKeys(crudVars[4]);
                //Thread.Sleep(500);
                ////AGREGAR REGISTRO
                //var ElementList = PruebaCRUD.NavClass(desktopSession);
                //desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 215, 15);
                //desktopSession.Mouse.Click(null);
                Thread.Sleep(5000);
                //newFunctions_4.ScreenshotNuevo("Agregar registro", true, file);
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Error de Appium: {ex.ToString()}");
            }
            return errorMessages;
        }
        public static List<string> ProcesoKNmLices(WindowsDriver<WindowsElement> desktopSession, List<string> crudVars, string file, int cont)
        {
            List<string> errorMessages = new List<string>();
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();
            try
            {
                Thread.Sleep(500);
                var TEdit = desktopSession.FindElementsByClassName("TEdit");
                if (cont == 0) 
                {
                    var Checkbx = desktopSession.FindElementsByName("Simulación");
                    Checkbx[0].Click();
                }

                Thread.Sleep(500);
                switch (crudVars[0])
                {
                    case ("Parcial"):
                        var TipoSol = desktopSession.FindElementsByName(crudVars[0]);
                        TipoSol[0].Click();
                        TipoSol[0].Click();
                        break;
                    case ("Fin de Año"):
                        TipoSol = desktopSession.FindElementsByName(crudVars[0]);
                        TipoSol[0].Click();
                        TipoSol[0].Click();
                        break;
                    case ("Definitiva"):
                        TipoSol = desktopSession.FindElementsByName(crudVars[0]);
                        TipoSol[0].Click();
                        TipoSol[0].Click();
                        break;
                    case ("Consolidada"):
                        TipoSol = desktopSession.FindElementsByName(crudVars[0]);
                        TipoSol[0].Click();
                        TipoSol[0].Click();
                        break;
                    case ("Sustitución Patronal"):
                        TipoSol = desktopSession.FindElementsByName(crudVars[0]);
                        TipoSol[0].Click();
                        TipoSol[0].Click();
                        break;
                    case ("Parcial Fondo"):
                        TipoSol = desktopSession.FindElementsByName(crudVars[0]);
                        TipoSol[0].Click();
                        TipoSol[0].Click();
                        break;
                }
                //Fecha Pago Liquidación
                TEdit[2].SendKeys(crudVars[1]);
                TEdit[2].SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                //Fecha Pago Cesantias
                TEdit[1].SendKeys(crudVars[2]);
                TEdit[1].SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                //Fecha Pago Intereses
                TEdit[0].SendKeys(crudVars[3]);
                TEdit[0].SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);

            }
            catch (Exception ex)
            {
                errorMessages.Add($"Error de Appium: {ex.ToString()}");
            }
            return errorMessages;
        }
        public static List<string> VentanasConfirmacionAceptar(WindowsDriver<WindowsElement> desktopSession, string file)
        {
            List<string> errorMessages = new List<string>();
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();
            try 
            { 
			    WindowsDriver<WindowsElement> rootSession = null;
			    Thread.Sleep(100);
			    rootSession = PruebaCRUD.RootSession();
			    rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
			    var allFrame = rootSession.FindElementsByClassName("IEFrame");
			    Thread.Sleep(100);
			    new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) + 120, (allFrame[0].Size.Height / 2) + 90).DoubleClick().Perform();
			    Thread.Sleep(5000);
			    newFunctions_4.ScreenshotNuevo("Fin del Proceso", true, file);
			    //Thread.Sleep(500);
			    //new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) + 110, (allFrame[0].Size.Height / 2) + 90).DoubleClick().Perform();
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Error de Appium: {ex.ToString()}");
            }
            return errorMessages;
        }
        public static List<string> VentanasConfirmacionSi(WindowsDriver<WindowsElement> desktopSession, string file)
        {
            List<string> errorMessages = new List<string>();
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();
            try
            {
                WindowsDriver<WindowsElement> rootSession = null;
                Thread.Sleep(100);
                rootSession = PruebaCRUD.RootSession();
                rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
                var allFrame = rootSession.FindElementsByClassName("IEFrame");
                Thread.Sleep(100);
                new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) + 60, (allFrame[0].Size.Height / 2) + 80).DoubleClick().Perform();
                Thread.Sleep(5000);
                newFunctions_4.ScreenshotNuevo("Fin del Proceso", true, file);
                //Thread.Sleep(500);
                //new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) + 110, (allFrame[0].Size.Height / 2) + 90).DoubleClick().Perform();
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Error de Appium: {ex.ToString()}");
            }
            return errorMessages;
        }
        public static List<string> QBEQry1(WindowsDriver<WindowsElement> desktopSession, string bandera, string QbeFilter, string file, string QbeFilter2, string campo = null, string campo2 = null)
        {
            List<string> errorMessages = new List<string>();
            if (bandera == "0")
            {
                var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
                {
                    Timeout = TimeSpan.FromSeconds(120),
                    PollingInterval = TimeSpan.FromSeconds(1)
                };
                wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
                wait.Until(driver =>
                {
                    WindowsDriver<WindowsElement> rootSession = null;
                    List<Point> coordenadas = FuncionesGlobales.coordinatesFinder(desktopSession, 244, 202, 32);

                    int coordX = coordenadas[0].X;
                    int coordY = coordenadas[0].Y;
                    var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));

                    bool IsDisplayedQbe = false;


                    desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordX, coordY);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    newFunctions_4.ScreenshotNuevo("QBE", true, file);
                    rootSession = RootSession();
                    rootSession = ReloadQbeSession(rootSession, "TKQBE");
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
                        if (!string.IsNullOrEmpty(QbeFilter))
                        {
                            elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/Pane[contains(@ClassName, 'TStringGrid')]/*"));
                            if (campo != null)
                            {
                                elements[0].Click();
                                int m = Convert.ToInt32(campo);
                                for (int i = 1; i <= m; i++)
                                {
                                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                                }
                                rootSession.Keyboard.SendKeys(QbeFilter);

                                Thread.Sleep(1000);
                                int m2 = Convert.ToInt32(campo2);
                                int ValIdentificacion = m2 - m;
                                for (int i = 1; i <= ValIdentificacion; i++)
                                {
                                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                                }
                                Thread.Sleep(1000);
                                for (int i = 1; i <= 10; i++) { rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace); }
                                rootSession.Keyboard.SendKeys(QbeFilter2);
                                Thread.Sleep(1000);
                                //elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                                //rootSession.Mouse.MouseMove(elements[8].Coordinates, 40, 80);
                                //rootSession.Mouse.Click(null);
                                Thread.Sleep(500);
                            }
                            else
                            {
                                elements[0].Click();
                                int m = Convert.ToInt32(campo);
                                for (int i = 1; i <= m; i++)
                                {
                                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                                }
                                rootSession.Keyboard.SendKeys(QbeFilter);

                                Thread.Sleep(1000);
                                int m2 = Convert.ToInt32(campo2);
                                int ValIdentificacion = m2 - m;
                                for (int i = 1; i <= ValIdentificacion; i++)
                                {
                                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                                }
                                Thread.Sleep(1000);
                                for (int i = 1; i <= 10; i++) { rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace); }
                                rootSession.Keyboard.SendKeys(QbeFilter2);
                                Thread.Sleep(1000);
                                //elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                                //rootSession.Mouse.MouseMove(elements[8].Coordinates, 40, 80);
                                //rootSession.Mouse.Click(null);
                                Thread.Sleep(500);
                            }
                        }
                        //Clic en Aceptar
                        /*var Element = rootSession.FindElementByClassName("TTabSheet");
                        rootSession.Mouse.MouseMove(Element.Coordinates, 465, 295);
                        rootSession.Mouse.Click(null);
                        Thread.Sleep(500);*/
                        elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                        rootSession.Mouse.MouseMove(elements[8].Coordinates, 20, 20);
                        rootSession.Mouse.Click(null);
                        Thread.Sleep(500);

                        if (string.IsNullOrEmpty(QbeFilter))
                        {
                            rootSession = ReloadQbeSession(rootSession, "TKQBE");
                            elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("//Button[contains(@Name, 'Sí')]"));
                            rootSession.Mouse.MouseMove(elements[0].Coordinates, 20, 20);
                            rootSession.Mouse.Click(null);
                        }
                        Thread.Sleep(500);
                        newFunctions_4.ScreenshotNuevo("Registros QBE", true, file);
                    }

                    return rootSession != null;
                });
            }
            else
            {
                var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
                {
                    Timeout = TimeSpan.FromSeconds(120),
                    PollingInterval = TimeSpan.FromSeconds(1)
                };
                wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
                wait.Until(driver =>
                {
                    WindowsDriver<WindowsElement> rootSession = null;
                    List<Point> coordenadas = FuncionesGlobales.coordinatesFinder(desktopSession, 224, 209, 6);

                    int coordX = coordenadas[0].X;
                    int coordY = coordenadas[0].Y;
                    var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));

                    bool IsDisplayedQbe = false;


                    desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordX, coordY);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    newFunctions_4.ScreenshotNuevo("QBE", true, file);
                    rootSession = RootSession();
                    rootSession = FuncionesGlobales.ReloadQbeSession(rootSession, "TKQBE");
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
                        if (!string.IsNullOrEmpty(QbeFilter))
                        {
                            elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/Pane[contains(@ClassName, 'TStringGrid')]/*"));
                            if (campo != null)
                            {
                                elements[0].Click();
                                int m = Convert.ToInt32(campo);
                                for (int i = 1; i <= m; i++)
                                {
                                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                                }
                                rootSession.Keyboard.SendKeys(QbeFilter);

                                Thread.Sleep(1000);
                                int m2 = Convert.ToInt32(campo2);
                                int ValIdentificacion = m2 - m;
                                for (int i = 1; i <= ValIdentificacion; i++)
                                {
                                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                                }
                                Thread.Sleep(1000);
                                rootSession.Keyboard.SendKeys(QbeFilter2);
                                Thread.Sleep(1000);
                                //elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                                //rootSession.Mouse.MouseMove(elements[8].Coordinates, 40, 80);
                                //rootSession.Mouse.Click(null);
                                Thread.Sleep(500);
                            }
                            else
                            {
                                elements[0].Click();
                                int m = Convert.ToInt32(campo);
                                for (int i = 1; i <= m; i++)
                                {
                                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                                }
                                rootSession.Keyboard.SendKeys(QbeFilter);

                                Thread.Sleep(1000);
                                int m2 = Convert.ToInt32(campo2);
                                int ValIdentificacion = m2 - m;
                                for (int i = 1; i <= ValIdentificacion; i++)
                                {
                                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                                }
                                Thread.Sleep(1000);
                                rootSession.Keyboard.SendKeys(QbeFilter2);
                                Thread.Sleep(1000);
                                //elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                                //rootSession.Mouse.MouseMove(elements[8].Coordinates, 40, 80);
                                //rootSession.Mouse.Click(null);
                                Thread.Sleep(500);
                            }

                        }
                        /*//Clic en Aceptar
                        var Element = rootSession.FindElementByClassName("TTabSheet");
                        rootSession.Mouse.MouseMove(Element.Coordinates, 465, 295);
                        rootSession.Mouse.Click(null);
                        Thread.Sleep(500);*/
                        elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                        rootSession.Mouse.MouseMove(elements[8].Coordinates, 20, 20);
                        rootSession.Mouse.Click(null);
                        Thread.Sleep(500);

                        if (string.IsNullOrEmpty(QbeFilter))
                        {
                            rootSession = FuncionesGlobales.ReloadQbeSession(rootSession, "TKQBE");
                            elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]"))[0].FindElements(By.XPath("//Button[contains(@Name, 'Sí')]"));
                            rootSession.Mouse.MouseMove(elements[0].Coordinates, 20, 20);
                            rootSession.Mouse.Click(null);
                        }
                        Thread.Sleep(500);
                        newFunctions_4.ScreenshotNuevo("Registros QBE", true, file);
                    }

                    return rootSession != null;
                });
            }
            return errorMessages;
        }
    }
}

   
