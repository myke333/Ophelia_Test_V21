using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
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

namespace Ophelia_Test_V21.AFuncionesGlobales.FuncionesProcesos.FuncionesProcesoKNmLivac
{
    class FuncionesKNmLivac : FuncionesVitales
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

        public static List<string> ProcesoKNmLivac(WindowsDriver<WindowsElement> desktopSession, List<string> crudVars, string file, int cont)
        {
            List<string> errorMessages = new List<string>();
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();

            try
            {
                //Tipo de liquidacion/ proceso de vacaciones
                var TipoLiq = desktopSession.FindElementsByName(crudVars[6]);
                Thread.Sleep(500);
                TipoLiq[0].Click();
                Thread.Sleep(1000);

                //Fecha de liquidacion
                var ElementList = desktopSession.FindElementsByClassName("TEdit");
                ElementList[0].SendKeys(crudVars[5]);
                Thread.Sleep(1500);
                

                //Empleados A Liquidar
                if (cont > 0)
                {
                    var Empleados = desktopSession.FindElementsByName("Empleados A Liquidar");
                    Thread.Sleep(500);
                    Empleados[0].Click();
                }
                               
                //WindowsDriver<WindowsElement> rootSession = null;
                //Thread.Sleep(100);
                //rootSession = PruebaCRUD.RootSession();
                //rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
                //var allFrame = rootSession.FindElementsByClassName("IEFrame");
                //Thread.Sleep(100);
                //new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) + 120, (allFrame[0].Size.Height / 2) + 90).Click().Perform();
                //Thread.Sleep(10000);
                //newFunctions_4.ScreenshotNuevo("Fin del Proceso", true, file);
                //Thread.Sleep(500);
                //new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) + 110, (allFrame[0].Size.Height / 2) + 90).DoubleClick().Perform();
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Error de Appium: {ex.ToString()}");
            }
            return errorMessages;
        }
        public static List<string> ProcesoKNmLivac2(WindowsDriver<WindowsElement> desktopSession, List<string> crudVars, string file, int cont)
        {
            List<string> errorMessages = new List<string>();
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();

            try
            {
                var AdicionarDatosQuery = desktopSession.FindElementsByName("Adicionar Datos Query");
                Thread.Sleep(500);
                AdicionarDatosQuery[0].Click();

                var ParametrosLiquidacion = desktopSession.FindElementsByName("Parametros Liquidación");
                Thread.Sleep(500);
                ParametrosLiquidacion[0].Click();
                Thread.Sleep(1000);

                if (crudVars[6] != "Consolidación ")
                {
                    //Tipo de Vacaciones
                    var TipoVacaciones = desktopSession.FindElementsByName(crudVars[4]);
                    Thread.Sleep(500);
                    TipoVacaciones[0].Click();
                    Thread.Sleep(1000);
                    //Proceso de campos de textos
                    var TEdit = desktopSession.FindElementsByClassName("TEdit");
                    Thread.Sleep(500);
                    //Se envia variable Dias Tiempo
                    TEdit[8].SendKeys(crudVars[0]);
                    Thread.Sleep(500);
                    //Se envia variable Dias Dinero
                    TEdit[7].SendKeys(crudVars[1]);
                    Thread.Sleep(500);
                    //Se envia variable Fecha Inicio
                    TEdit[6].SendKeys(crudVars[2]);
                    Thread.Sleep(500);
                    //Se envian variables Fecha Pago
                    TEdit[5].SendKeys(crudVars[3]);
                    Thread.Sleep(1000);
                }
                

                //Click Boton Aceptar
                var Aceptar = desktopSession.FindElementsByName("Aceptar");
                Thread.Sleep(500);
                Aceptar[0].Click();
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Error de Appium: {ex.ToString()}");
            }
            return errorMessages;
        }
        public static List<string> QBEQry(WindowsDriver<WindowsElement> desktopSession, string bandera, string QbeFilter, string file, string QbeFilter2, string campo = null, string campoid = null)
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
                            if (campo != null)
                            {
                                //Selecciono ventana consultas
                                rootSession.FindElementByName("Consultas").Click();
                                Thread.Sleep(2000);
                                //Selecciono opción de biodata
                                rootSession.FindElementByName("Biodata").Click();
                                Thread.Sleep(2000);
                                //Doy cLic en realizar filtro
                                rootSession.FindElementByName("Realizar Filtros").Click();
                                Thread.Sleep(2000);
                                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                                Thread.Sleep(1000);
                                rootSession.Keyboard.SendKeys(QbeFilter);
                                //Doy cLic en general
                                rootSession.FindElementByName("General").Click();
                                int m = Convert.ToInt32(campo);
                                for (int i = 1; i <= m; i++)
                                {
                                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                                }
                                rootSession.Keyboard.SendKeys(QbeFilter2);
                                Thread.Sleep(1000);
                                int m2 = Convert.ToInt32(campoid);
                                int ValIdentificacion = m - m2;
                                for (int i = 1; i <= ValIdentificacion; i++)
                                {
                                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowUp);
                                }
                                Thread.Sleep(1000);
                                rootSession.FindElementByName("Asc").Click();
                                Thread.Sleep(1000);
                                elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                                rootSession.Mouse.MouseMove(elements[8].Coordinates, 40, 80);
                                rootSession.Mouse.Click(null);
                                Thread.Sleep(500);
                            }
                            else
                            {
                                //Selecciono ventana consultas
                                rootSession.FindElementByName("Consultas").Click();
                                Thread.Sleep(2000);
                                //Selecciono opción de biodata
                                rootSession.FindElementByName("Biodata").Click();
                                Thread.Sleep(2000);
                                //Doy cLic en realizar filtro
                                rootSession.FindElementByName("Realizar Filtros").Click();
                                Thread.Sleep(2000);
                                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                                Thread.Sleep(1000);
                                rootSession.Keyboard.SendKeys(QbeFilter);
                                //Doy cLic en general
                                rootSession.FindElementByName("General").Click();
                                int m = Convert.ToInt32(campo);
                                for (int i = 1; i <= m; i++)
                                {
                                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                                }
                                rootSession.Keyboard.SendKeys(QbeFilter2);
                                Thread.Sleep(1000);
                                int m2 = Convert.ToInt32(campoid);
                                int ValIdentificacion = m - m2;
                                for (int i = 1; i <= ValIdentificacion; i++)
                                {
                                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowUp);
                                }
                                Thread.Sleep(1000);
                                rootSession.FindElementByName("Asc").Click();
                                Thread.Sleep(1000);
                                elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                                rootSession.Mouse.MouseMove(elements[8].Coordinates, 40, 80);
                                rootSession.Mouse.Click(null);
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
                            if (campo != null)
                            {
                                //Selecciono ventana consultas
                                rootSession.FindElementByName("Consultas").Click();
                                Thread.Sleep(2000);
                                //Selecciono opción de biodata
                                rootSession.FindElementByName("Biodata").Click();
                                Thread.Sleep(2000);
                                //Doy cLic en realizar filtro
                                rootSession.FindElementByName("Realizar Filtros").Click();
                                Thread.Sleep(2000);
                                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                                Thread.Sleep(1000);
                                rootSession.Keyboard.SendKeys(QbeFilter);
                                //Doy cLic en general
                                rootSession.FindElementByName("General").Click(); 
                                int m = Convert.ToInt32(campo);
                                for (int i = 1; i <= m; i++)
                                {
                                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                                }
                                rootSession.Keyboard.SendKeys(QbeFilter2);
                                int m2 = Convert.ToInt32(campoid);
                                int ValIdentificacion = m - m2;
                                for (int i = 0; i <= ValIdentificacion; i++)
                                {
                                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowUp);
                                }
                            }
                            else
                            {
                                //Selecciono ventana consultas
                                rootSession.FindElementByName("Consultas").Click();
                                Thread.Sleep(2000);
                                //Selecciono opción de biodata
                                rootSession.FindElementByName("Biodata").Click();
                                Thread.Sleep(2000);
                                //Doy cLic en realizar filtro
                                rootSession.FindElementByName("Realizar Filtros").Click();
                                Thread.Sleep(2000);
                                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                                Thread.Sleep(1000); //General
                                rootSession.Keyboard.SendKeys(QbeFilter);
                                Thread.Sleep(2000);
                                //Doy cLic en general
                                rootSession.FindElementByName("General").Click(); 
                                int m = Convert.ToInt32(campo);
                                for (int i = 1; i <= m; i++)
                                {
                                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                                }
                                rootSession.Keyboard.SendKeys(QbeFilter2);
                                Thread.Sleep(1000);
                                rootSession.FindElementByName("Asc").Click();
                                Thread.Sleep(1000); 
                                elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                                rootSession.Mouse.MouseMove(elements[8].Coordinates, 40, 80);
                                rootSession.Mouse.Click(null);
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

        public static List<string> QBEQrylices(WindowsDriver<WindowsElement> desktopSession, string bandera, string QbeFilter, string file, string QbeFilter2, string campo = null, string campoid = null)
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
                            if (campo != null)
                            {
                                //Selecciono ventana consultas
                                rootSession.FindElementByName("Consultas").Click();
                                Thread.Sleep(2000);
                                //Selecciono opción de biodata
                                rootSession.FindElementByName("Biodata").Click();
                                Thread.Sleep(2000);
                                //Doy cLic en realizar filtro
                                rootSession.FindElementByName("Realizar Filtros").Click();
                                Thread.Sleep(2000);
                                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                                Thread.Sleep(1000);
                                rootSession.Keyboard.SendKeys(QbeFilter);
                                //Doy cLic en general
                                rootSession.FindElementByName("General").Click();
                                int m = Convert.ToInt32(campo);
                                for (int i = 1; i <= m; i++)
                                {
                                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                                }
                                rootSession.Keyboard.SendKeys(QbeFilter2);
                                Thread.Sleep(1000);
                                int m2 = Convert.ToInt32(campoid);
                                int ValIdentificacion = m - m2;
                                for (int i = 1; i <= ValIdentificacion; i++)
                                {
                                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowUp);
                                }
                                Thread.Sleep(1000);
                                rootSession.FindElementByName("Asc").Click();
                                Thread.Sleep(1000);
                                elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                                rootSession.Mouse.MouseMove(elements[8].Coordinates, 40, 80);
                                rootSession.Mouse.Click(null);
                                Thread.Sleep(500);
                            }
                            else
                            {
                                //Selecciono ventana consultas
                                rootSession.FindElementByName("Consultas").Click();
                                Thread.Sleep(2000);
                                //Selecciono opción de biodata
                                rootSession.FindElementByName("Biodata").Click();
                                Thread.Sleep(2000);
                                //Doy cLic en realizar filtro
                                rootSession.FindElementByName("Realizar Filtros").Click();
                                Thread.Sleep(2000);
                                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                                Thread.Sleep(1000);
                                rootSession.Keyboard.SendKeys(QbeFilter);
                                //Doy cLic en general
                                rootSession.FindElementByName("General").Click();
                                int m = Convert.ToInt32(campo);
                                for (int i = 1; i <= m; i++)
                                {
                                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                                }
                                rootSession.Keyboard.SendKeys(QbeFilter2);
                                Thread.Sleep(1000);
                                int m2 = Convert.ToInt32(campoid);
                                int ValIdentificacion = m - m2;
                                for (int i = 1; i <= ValIdentificacion; i++)
                                {
                                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowUp);
                                }
                                Thread.Sleep(1000);
                                rootSession.FindElementByName("Asc").Click();
                                Thread.Sleep(1000);
                                elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                                rootSession.Mouse.MouseMove(elements[8].Coordinates, 40, 80);
                                rootSession.Mouse.Click(null);
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
                            if (campo != null)
                            {
                                //Selecciono ventana consultas
                                rootSession.FindElementByName("Consultas").Click();
                                Thread.Sleep(2000);
                                //Selecciono opción de biodata
                                rootSession.FindElementByName("Biodata").Click();
                                Thread.Sleep(2000);
                                //Doy cLic en realizar filtro
                                rootSession.FindElementByName("Realizar Filtros").Click();
                                Thread.Sleep(2000);
                                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                                Thread.Sleep(1000);
                                rootSession.Keyboard.SendKeys(QbeFilter);
                                //Doy cLic en general
                                rootSession.FindElementByName("General").Click();
                                int m = Convert.ToInt32(campo);
                                for (int i = 1; i <= m; i++)
                                {
                                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                                }
                                rootSession.Keyboard.SendKeys(QbeFilter2);
                                int m2 = Convert.ToInt32(campoid);
                                int ValIdentificacion = m - m2;
                                for (int i = 0; i <= ValIdentificacion; i++)
                                {
                                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowUp);
                                }
                            }
                            else
                            {
                                //Selecciono ventana consultas
                                rootSession.FindElementByName("Consultas").Click();
                                Thread.Sleep(2000);
                                //Selecciono opción de biodata
                                rootSession.FindElementByName("Biodata").Click();
                                Thread.Sleep(2000);
                                //Doy cLic en realizar filtro
                                rootSession.FindElementByName("Realizar Filtros").Click();
                                Thread.Sleep(2000);
                                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                                Thread.Sleep(1000); //General
                                rootSession.Keyboard.SendKeys(QbeFilter);
                                Thread.Sleep(2000);
                                //Doy cLic en general
                                rootSession.FindElementByName("General").Click();
                                int m = Convert.ToInt32(campo);
                                for (int i = 1; i <= m; i++)
                                {
                                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                                }
                                rootSession.Keyboard.SendKeys(QbeFilter2);
                                Thread.Sleep(1000);
                                rootSession.FindElementByName("Asc").Click();
                                Thread.Sleep(1000);
                                elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                                rootSession.Mouse.MouseMove(elements[8].Coordinates, 40, 80);
                                rootSession.Mouse.Click(null);
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
                                for (int i = 1; i <= 10; i++) { rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace); }
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
        public static List<string> ValidacionesKNmDiasn(WindowsDriver<WindowsElement> desktopSession, string IdInicial, string IdFinal, int cont, string Concepto, string FechaDesde, string FechaHasta)
        {
            List<string> errorMessages = new List<string>();
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();
            int IdInicialInt = Convert.ToInt32(IdInicial);
            int IdFinalInt = Convert.ToInt32(IdFinal);
            //int ConceptoInt = Convert.ToInt32(Concepto);
            bool IsDisplayedValue = false;
            try
            {
                var TDBGridInplaceEdit = desktopSession.FindElementsByClassName("TDBGridInplaceEdit");
                TDBGridInplaceEdit[0].Click();
                Thread.Sleep(1000);
                //Valida la posición donde arranca
                if (cont == 0)
                {
                    for (int i = 0; i < 500; i++)
                    {
                        string Identificacion = TDBGridInplaceEdit[0].Text;
                        int IdentificacionInt = Convert.ToInt32(Identificacion);
                        if (IdentificacionInt >= IdInicialInt && IdentificacionInt <= IdFinalInt)
                        {
                            IsDisplayedValue = true;
                            break;
                        }
                        else
                        {
                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                            Thread.Sleep(2000);
                        }
                    }
                }
                if (cont != 0)
                {
                    IsDisplayedValue = true;
                }
                if (IsDisplayedValue)
                {
                    if (cont != 0)
                    {
                        Thread.Sleep(2000);
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                        Thread.Sleep(500);
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Home);
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Home);
                    }
                    string IdentificacionReal = TDBGridInplaceEdit[0].Text;
                    int IdentificacionRealInt = Convert.ToInt32(IdentificacionReal);
                    if (IdentificacionRealInt >= IdInicialInt && IdentificacionRealInt <= IdFinalInt)
                    {
                        //Validacion Concepto Grilla
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(1000);
                        string ConceptObt = TDBGridInplaceEdit[0].Text;
                        if (ConceptObt != Concepto)
                        {
                            errorMessages.Add($"El Concepto del usuario {IdentificacionReal} esperado es: {Concepto} y el encontrado es: {ConceptObt}");
                        }
                        Thread.Sleep(2000);

                        //Validacion FechaDesdeEsp Grilla
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(1000);
                        string FechaDesdeObt = TDBGridInplaceEdit[0].Text;
                        Thread.Sleep(1000);
                        if (FechaDesde != FechaDesdeObt)
                        {
                            errorMessages.Add($"La FechaDesde del usuario {IdentificacionReal} esperada es: {FechaDesde} y la encontrada es: {FechaDesdeObt}");
                        }
                        Thread.Sleep(1000);
                        //Validacion FechaHastaEsp Grilla
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(1000);
                        string FechaHastaObt = TDBGridInplaceEdit[0].Text;
                        Thread.Sleep(2000);
                        if (FechaHasta != FechaHastaObt)
                        {
                            errorMessages.Add($"La FechaHasta del usuario {IdentificacionReal} esperada es: {FechaHasta} y la encontrada es: {FechaHastaObt}");
                        }
                        Thread.Sleep(1000);
                    }
                    else
                    {
                        errorMessages.Add($"Valor no es igual o posiblemente esta vacio");
                    }
                }
                else
                {
                    errorMessages.Add($"No puede encontrar valor dentro de la grilla entre ese rango de identificación");
                }
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Error de Appium: {ex.ToString()}");
            }
            return errorMessages;

        }

        public static List<string> ValidacionesKNmVacac(WindowsDriver<WindowsElement> desktopSession, string codProgram, string IdInicial, string IdFinal, int cont, string Concepto, string FechaDesde, string FechaHasta, string FechaICausacion, string FechaFCausacion, string SueldoBase)
        {
            List<string> errorMessages = new List<string>();
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();
            int IdInicialInt = Convert.ToInt32(IdInicial);
            int IdFinalInt = Convert.ToInt32(IdFinal);
            //int ConceptoInt = Convert.ToInt32(Concepto);
            bool IsDisplayedValue = false;
            try
            {
                var TDBGridInplaceEdit = desktopSession.FindElementsByClassName("TDBGridInplaceEdit");
                TDBGridInplaceEdit[0].Click();
                Thread.Sleep(1000);
                //Valida la posición donde arranca
                if (cont == 0)
                {
                    for (int i = 0; i < 500; i++)
                    {
                        string Identificacion = TDBGridInplaceEdit[0].Text;
                        int IdentificacionInt = Convert.ToInt32(Identificacion);
                        if (IdentificacionInt >= IdInicialInt && IdentificacionInt <= IdFinalInt)
                        {
                            IsDisplayedValue = true;
                            break;
                        }
                        else
                        {
                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                            Thread.Sleep(2000);
                        }
                    }
                }
                if (cont != 0)
                {
                    IsDisplayedValue = true;
                }
                if (IsDisplayedValue)
                {
                    if (cont != 0)
                    {
                        Thread.Sleep(2000);
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                        Thread.Sleep(500);
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Home);
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Home);
                    }
                    string IdentificacionReal = TDBGridInplaceEdit[0].Text;
                    int IdentificacionRealInt = Convert.ToInt32(IdentificacionReal);
                    if (IdentificacionRealInt >= IdInicialInt && IdentificacionRealInt <= IdFinalInt)
                    {
                        if (codProgram == "KNmDiasn")
                        {
                            //Validacion Concepto Grilla
                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                            Thread.Sleep(1000);
                            string ConceptObt = TDBGridInplaceEdit[0].Text;
                            if (ConceptObt != Concepto)
                            {
                                errorMessages.Add($"El Concepto del usuario {IdentificacionReal} esperado es: {Concepto} y el encontrado es: {ConceptObt}");
                            }
                            Thread.Sleep(2000);

                            //Validacion FechaDesdeEsp Grilla
                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                            Thread.Sleep(1000);
                            string FechaDesdeObt = TDBGridInplaceEdit[0].Text;
                            Thread.Sleep(1000);
                            if (FechaDesde != FechaDesdeObt)
                            {
                                errorMessages.Add($"La FechaDesde del usuario {IdentificacionReal} esperada es: {FechaDesde} y la encontrada es: {FechaDesdeObt}");
                            }
                            Thread.Sleep(1000);
                            //Validacion FechaHastaEsp Grilla
                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                            Thread.Sleep(1000);
                            string FechaHastaObt = TDBGridInplaceEdit[0].Text;
                            Thread.Sleep(2000);
                            if (FechaHasta != FechaHastaObt)
                            {
                                errorMessages.Add($"La FechaHasta del usuario {IdentificacionReal} esperada es: {FechaHasta} y la encontrada es: {FechaHastaObt}");
                            }
                            Thread.Sleep(1000);
                        }
                        if (codProgram == "KNmVacac")
                        {
                            //Validacion Fecha inicial causacion Grilla
                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                            Thread.Sleep(1000);
                            string FechaICausacionobt = TDBGridInplaceEdit[0].Text;
                            if (FechaICausacionobt != FechaICausacion)
                            {
                                errorMessages.Add($"La fecha Inicio Causacion del usuario {IdentificacionReal} esperado es: {FechaICausacion} y el encontrado es: {FechaICausacionobt}");
                            }
                            Thread.Sleep(2000);

                            //Validacion Fecha Final causacion Grilla
                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                            Thread.Sleep(1000);
                            string FechafCausacionobt = TDBGridInplaceEdit[0].Text;
                            Thread.Sleep(1000);
                            if (FechafCausacionobt != FechaFCausacion)
                            {
                                errorMessages.Add($"La fecha Final Causacion del usuario {IdentificacionReal} esperada es: {FechaFCausacion} y la encontrada es: {FechafCausacionobt}");
                            }
                            Thread.Sleep(1000);
                            //Validacion FechaDesdeEsp Grilla
                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                            Thread.Sleep(1000);
                            string FechaDesdeObt = TDBGridInplaceEdit[0].Text;
                            Thread.Sleep(1000);
                            if (FechaDesde != FechaDesdeObt)
                            {
                                errorMessages.Add($"La FechaDesde del usuario {IdentificacionReal} esperada es: {FechaDesde} y la encontrada es: {FechaDesdeObt}");
                            }
                            Thread.Sleep(1000);
                            //Validacion FechaHastaEsp Grilla
                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                            Thread.Sleep(1000);
                            string FechaHastaObt = TDBGridInplaceEdit[0].Text;
                            Thread.Sleep(2000);
                            if (FechaHasta != FechaHastaObt)
                            {
                                errorMessages.Add($"La FechaHasta del usuario {IdentificacionReal} esperada es: {FechaHasta} y la encontrada es: {FechaHastaObt}");
                            }
                            Thread.Sleep(1000);
                            //Validacion Sueldo Base
                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                            Thread.Sleep(1000);
                            string SueldoBaseObt = TDBGridInplaceEdit[0].Text;
                            Thread.Sleep(2000);
                            if (SueldoBase != SueldoBaseObt)
                            {
                                errorMessages.Add($"El sueldo Base del usuario {IdentificacionReal} esperada es: {SueldoBase} y la encontrada es: {SueldoBaseObt}");
                            }
                            Thread.Sleep(1000);
                        }
                        if (codProgram == "KNmAusen")
                        {
                            //Validacion Concepto Grilla
                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                            Thread.Sleep(1000);
                            string ConceptObt = TDBGridInplaceEdit[0].Text;
                            if (ConceptObt != Concepto)
                            {
                                errorMessages.Add($"El Concepto del usuario {IdentificacionReal} esperado es: {Concepto} y el encontrado es: {ConceptObt}");
                            }
                            Thread.Sleep(2000);
                            //Validacion FechaDesdeEsp Grilla
                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                            Thread.Sleep(1000);
                            string FechaDesdeObt = TDBGridInplaceEdit[0].Text;
                            Thread.Sleep(1000);
                            if (FechaDesde != FechaDesdeObt)
                            {
                                errorMessages.Add($"La FechaDesde del usuario {IdentificacionReal} esperada es: {FechaDesde} y la encontrada es: {FechaDesdeObt}");
                            }
                            Thread.Sleep(1000);
                            //Validacion FechaHastaEsp Grilla
                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                            Thread.Sleep(1000);
                            string FechaHastaObt = TDBGridInplaceEdit[0].Text;
                            Thread.Sleep(2000);
                            if (FechaHasta != FechaHastaObt)
                            {
                                errorMessages.Add($"La FechaHasta del usuario {IdentificacionReal} esperada es: {FechaHasta} y la encontrada es: {FechaHastaObt}");
                            }
                            Thread.Sleep(1000);
                        }
                        if (codProgram == "KNmPreno")
                        {
                            //Validacion Concepto Grilla
                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                            Thread.Sleep(1000);
                            string ConceptObt = TDBGridInplaceEdit[0].Text;
                            if (ConceptObt != Concepto)
                            {
                                errorMessages.Add($"El Concepto del usuario {IdentificacionReal} esperado es: {Concepto} y el encontrado es: {ConceptObt}");
                            }
                            Thread.Sleep(2000);
                            //Validacion FechaDesdeEsp Grilla
                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                            Thread.Sleep(1000);
                            string FechaDesdeObt = TDBGridInplaceEdit[0].Text;
                            Thread.Sleep(1000);
                            if (FechaDesde != FechaDesdeObt)
                            {
                                errorMessages.Add($"La FechaDesde del usuario {IdentificacionReal} esperada es: {FechaDesde} y la encontrada es: {FechaDesdeObt}");
                            }
                            Thread.Sleep(1000);
                            //Validacion FechaHastaEsp Grilla
                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                            Thread.Sleep(1000);
                            string FechaHastaObt = TDBGridInplaceEdit[0].Text;
                            Thread.Sleep(2000);
                            if (FechaHasta != FechaHastaObt)
                            {
                                errorMessages.Add($"La FechaHasta del usuario {IdentificacionReal} esperada es: {FechaHasta} y la encontrada es: {FechaHastaObt}");
                            }
                            Thread.Sleep(1000);
                        }

                    }
                    else
                    {
                        errorMessages.Add($"Valor no es igual o posiblemente esta vacio");
                    }
                }
                else
                {
                    errorMessages.Add($"No puede encontrar valor dentro de la grilla entre ese rango de identificación");
                }
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Error de Appium: {ex.ToString()}");
            }
            return errorMessages;

        }


    }
}

