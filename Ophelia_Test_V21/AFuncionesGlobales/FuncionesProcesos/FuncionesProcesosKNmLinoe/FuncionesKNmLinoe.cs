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

namespace Ophelia_Test_V21.AFuncionesGlobales.FuncionesProcesos.FuncionesProcesosKNmLinoe
{
    class FuncionesKNmLinoe : FuncionesVitales
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
        public static List<string> TablasDetallles(WindowsDriver<WindowsElement> desktopSession, string file, string NomPestana, string NomTablaDetalle, int j, int Secuencial, int CamposTabla)
        {
            List<string> errorMessages = new List<string>();
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();
            try
            {
                if (Secuencial > 0) 
                {
                    var Pestana = desktopSession.FindElementsByName(NomPestana);
                    Pestana[0].Click();
                    var TablaDetalle = desktopSession.FindElementsByName(NomTablaDetalle);
                    TablaDetalle[0].Click();

                    WindowsDriver<WindowsElement> rootSession = null;
                    Thread.Sleep(100);
                    rootSession = PruebaCRUD.RootSession();
                    rootSession = PruebaCRUD.ReloadSession(rootSession, "TDBGrid");

                    var Cajas= desktopSession.FindElementsByClassName("TDBGrid");
                    Cajas[j].Click();

                    //Funciones para encontrar secuencial 1
                    for (int i = 0; i < Secuencial; i++) {desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowLeft);}
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    var celdaIncap = desktopSession.FindElementsByClassName("TDBGridInplaceEdit");
                    while (celdaIncap[1].Text != "1") 
                    {
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowUp);
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    }
                    //////
                    //Llenado de datos
                    List<string> Incapacidad = new List<string>();
                    Incapacidad.Add(celdaIncap[1].Text);
                    for (int i = 0; i < CamposTabla; i++)
                    {
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        Incapacidad.Add(celdaIncap[1].Text);
                    }
                }
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Error de Appium: {ex.ToString()}");
            }
            return errorMessages;
        }
        public static List<string> QBEQry(WindowsDriver<WindowsElement> desktopSession, string bandera, string QbeFilter, string QbeFilter2, string QbeFilter3, string file, string campo = null, string campo2 = null)
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
                                //for (int i = 1; i <= 10; i++) { rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace); }
                                
                                int m = Convert.ToInt32(campo);
                                for (int i = 1; i <= m; i++){rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);}
                                rootSession.Keyboard.SendKeys(QbeFilter);
                                Thread.Sleep(1000);

                                int m2 = Convert.ToInt32(campo2);
                                int ValIdentificacion = m - m2;
                                elements[0].Click();
                                for (int i = 1; i <= ValIdentificacion; i++){rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowUp); }
                                Thread.Sleep(1000);
                                //for (int i = 1; i <= 10; i++) { rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace); }
                                rootSession.Keyboard.SendKeys(QbeFilter2);
                                Thread.Sleep(1000);
                                for (int i = 1; i <= m2; i++) {rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowUp); }
                                rootSession.Keyboard.SendKeys(QbeFilter3);
                                rootSession.FindElementByName("Asc").Click();
                            }
                            else
                            {
                                errorMessages.Add($"No existe valor en Campo");
                            } 
                        }
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
            } //3 FILTROS IDENTIFICACION POSICION 1
            else if(bandera == "5") //3 FILTROS IDENTIFICACION POSICION 2
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
                                //for (int i = 1; i <= 10; i++) { rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace); }

                                int m = Convert.ToInt32(campo);
                                for (int i = 1; i <= m; i++) { rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown); }
                                rootSession.Keyboard.SendKeys(QbeFilter);
                                Thread.Sleep(1000);

                                int m2 = Convert.ToInt32(campo2);
                                int ValIdentificacion = m - m2;
                                elements[0].Click();
                                for (int i = 1; i <= ValIdentificacion; i++) { rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowUp); }
                                Thread.Sleep(1000);
                                //for (int i = 1; i <= 10; i++) { rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace); }
                                rootSession.Keyboard.SendKeys(QbeFilter2);
                                Thread.Sleep(1000);
                                for (int i = 1; i < m2; i++) { rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowUp); }
                                rootSession.Keyboard.SendKeys(QbeFilter3);
                                rootSession.FindElementByName("Asc").Click();
                            }
                            else
                            {
                                errorMessages.Add($"No existe valor en Campo");
                            }
                        }
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
                                Thread.Sleep(500);
                               // for (int i = 1; i <= 10; i++) { rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace); }
                                rootSession.Keyboard.SendKeys(QbeFilter2);
                                Thread.Sleep(1000);
                                //elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                                //rootSession.Mouse.MouseMove(elements[8].Coordinates, 40, 80);
                                //rootSession.Mouse.Click(null);
                                //Thread.Sleep(500);
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

            return errorMessages;
        }
        public static List<string> ProcesoKNmLinoe(WindowsDriver<WindowsElement> desktopSession, List<string> crudVars, string file, int cont)
        {
            List<string> errorMessages = new List<string>();
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();
            try
            {
                /// Obejeots Checkbox
                var CheckbxNormal = desktopSession.FindElementsByName("Normal");
                var CheckbxExtranómina = desktopSession.FindElementsByName("Extranómina");
                var CheckbxAdicional = desktopSession.FindElementsByName("Adicional");
                var CheckbxDefinitiva = desktopSession.FindElementsByName("Definitiva");
                var CheckbxCesantías = desktopSession.FindElementsByName("Cesantías");
                var CheckbxVacaciones = desktopSession.FindElementsByName("Vacaciones");
                var CheckbxPrimas = desktopSession.FindElementsByName("Primas");
                var CheckbxPrimaExtralegal = desktopSession.FindElementsByName("Prima Extralegal");
                Thread.Sleep(500);
                /// CheckBox Procesos
                var CheckbxProcesos = desktopSession.FindElementsByName(crudVars[0]);// General //Comparativo
                CheckbxProcesos[0].Click();
                /// CheckBox Tipo de liquidacion 
                var CheckbxTipoLiq = desktopSession.FindElementsByName(crudVars[1]);//Nómina Individual // Nómina Individual de Ajuste
                CheckbxTipoLiq[0].Click();
                /// CheckBox Tipo de Ajuste
                if (crudVars[1] == "Nómina Individual de Ajuste")
                {
                    var CheckbxTipoAdj = desktopSession.FindElementsByName(crudVars[2]);//Eliminar // Reemplazar
                    CheckbxTipoAdj[0].Click();
                }
                Thread.Sleep(500);

                /// Encontrar TEdit 
                var TEdit = desktopSession.FindElementsByClassName("TEdit");
                // PREFIJO 
                TEdit[2].SendKeys(crudVars[3]);
                // TRM
                TEdit[1].SendKeys(crudVars[4]);
                //CUNE TEdit = 0
                TEdit[0].SendKeys(crudVars[5]);
                //// Verificaciones 
                if (crudVars[7] != "1") { CheckbxNormal[0].Click(); }
                if (crudVars[8] != "1") { CheckbxExtranómina[0].Click(); }
                if (crudVars[9] != "1") { CheckbxAdicional[0].Click(); }
                if (crudVars[10] != "1") { CheckbxDefinitiva[0].Click(); }
                if (crudVars[11] != "1") { CheckbxCesantías[0].Click(); }
                if (crudVars[12] != "1") { CheckbxVacaciones[0].Click(); }
                if (crudVars[13] != "1") { CheckbxPrimas[0].Click(); }
                if (crudVars[14] != "1") { CheckbxPrimaExtralegal[0].Click(); }

                Thread.Sleep(500);
                //Fecha Inicial
                TEdit[5].SendKeys(crudVars[15]);
                TEdit[5].SendKeys(OpenQA.Selenium.Keys.Tab);
                //Fecha Final
                TEdit[4].SendKeys(crudVars[16]);
                TEdit[4].SendKeys(OpenQA.Selenium.Keys.Tab);
                //Fecha de Emision 
                TEdit[3].SendKeys(crudVars[17]);
                TEdit[3].SendKeys(OpenQA.Selenium.Keys.Tab);

                //TMemo = 0 observaciones TGroupBox
                var TMemo = desktopSession.FindElementsByClassName("TMemo");
                TMemo[0].SendKeys(crudVars[6]);

                //Tomar Fecha Pago
                var CheckbxFechaPago = desktopSession.FindElementsByName("Tomar Fecha Pago");
                if (crudVars[18] != "0") { CheckbxFechaPago[0].Click(); }
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Error de Appium: {ex.ToString()}");
            }
            return errorMessages;
        }

    }
}
