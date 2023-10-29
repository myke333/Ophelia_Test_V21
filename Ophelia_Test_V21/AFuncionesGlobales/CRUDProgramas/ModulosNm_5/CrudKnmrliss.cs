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


namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_5
{
    class CrudKnmrliss : FuncionesVitales
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

        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TEdit");
            ElementList[2].SendKeys(data[1]);
            ElementList[3].SendKeys(data[0]);
        }

        public static void ButtonDatos(WindowsDriver<WindowsElement> desktopSession)
        {
            List<string> errorMessages = new List<string>();
            desktopSession.FindElementByName("Datos").Click();
        }



        public static List<string> QBEQry(WindowsDriver<WindowsElement> desktopSession, string bandera, string QbeFilter, string file, string campo = null)
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
                    List<Point> coordenadas = FuncionesGlobales.coordinatesFinder(desktopSession, 227, 117, 0);

                    int coordX = coordenadas[0].X + 270;
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
                            }
                            else
                            {
                                elements[0].SendKeys(QbeFilter);
                            }
                        }
                        elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                        rootSession.Mouse.MouseMove(elements[8].Coordinates, 20, 20);
                        rootSession.Mouse.Click(null);
                        Thread.Sleep(500);

                        if (string.IsNullOrEmpty(QbeFilter))
                        {
                            rootSession = FuncionesGlobales.ReloadQbeSession(rootSession, "TKQBE");
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
                    List<Point> coordenadas = FuncionesGlobales.coordinatesFinder(desktopSession, 227, 117, 0);

                    int coordX = coordenadas[0].X + 265;
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
                            }
                            else
                            {
                                elements[0].SendKeys(QbeFilter);
                            }

                        }
                        elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                        rootSession.Mouse.MouseMove(elements[8].Coordinates, 20, 20);
                        rootSession.Mouse.Click(null);
                        Thread.Sleep(500);

                        if (string.IsNullOrEmpty(QbeFilter))
                        {
                            rootSession = FuncionesGlobales.ReloadQbeSession(rootSession, "TKQBE");
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



        public static void VentanaPreliminar(WindowsDriver<WindowsElement> desktopSession, string file, int tipo)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            if (tipo == 1)
            {
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                newFunctions_4.ScreenshotNuevo("Preliminar Detalles", true, file);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

            }
            else
            {
                for (int i = 0; i < 3; i++)
                {
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                }
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                newFunctions_4.ScreenshotNuevo("Preliminar Resumido", true, file);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

            }
        }

        public static void VentanaImprimir(WindowsDriver<WindowsElement> desktopSession, string file, int tipo)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            if (tipo == 1)
            {
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                newFunctions_4.ScreenshotNuevo("Imprimir Detalles", true, file);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

            }
            else
            {
                for (int i = 0; i < 3; i++)
                {
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                }
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                newFunctions_4.ScreenshotNuevo("Imprimir Resumido", true, file);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

            }


        }

        public static void ButtonImprimir(WindowsDriver<WindowsElement> desktopSession)
        {
            List<string> errorMessages = new List<string>();
            desktopSession.FindElementByName("Imprimir").Click();
        }

        public static void BotonImprimir(WindowsDriver<WindowsElement> desktopSession, string file)
        {

            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            rootSession = ReloadSession(rootSession, "TfrxPrintDialog");

            var scrollelement = rootSession.FindElementsByClassName("TComboBox");
            rootSession.Mouse.MouseMove(scrollelement[5].Coordinates, 10, 10);
            rootSession.Mouse.Click(null);
            rootSession.Mouse.Click(null);
            //newFunctions_4.ScreenshotNuevo("Ventana Impresión PDF", true, file);
            Thread.Sleep(1000);
            BotonImprimir2(desktopSession);

        }

        public static void BotonImprimir2(WindowsDriver<WindowsElement> desktopSession)
        {

            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            rootSession = ReloadSession(rootSession, "TfrxPrintDialog");
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
        }

    }
}
