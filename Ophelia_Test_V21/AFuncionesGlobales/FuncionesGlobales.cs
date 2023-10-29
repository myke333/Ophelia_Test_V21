using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Ophelia_Test_V21.AFuncionesGlobales;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using System.Data;
using System.CodeDom;
using System.CodeDom.Compiler;
using System.IO;
using System.Globalization;
using System.Configuration;
using System.Data.SqlClient;

namespace Ophelia_Test_V21
{
    class FuncionesGlobales : FuncionesVitales
    {
       static  public WindowsDriver<WindowsElement> RootSession()
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
                //Debugger.Launch();
                return ApplicationSession;
            }
            catch (Exception ex) { return null; }
        }

        public static void GetVersion(WindowsDriver<WindowsElement> desktopSession, string file)
        {
            WindowsDriver<WindowsElement> rootSession = null;
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));

            //FUNCION CERRAR VENTANA EMERGENTE EN EDGE
            string navigator = AbrirPrograma.SelectNavigator();
            if (navigator == "Edge")
            {
                //NOTIFICACION
                CerrarVentanaNotifiacion(desktopSession);
            }

            WindowsElement versionButtom = null;
            //
            wait.Until(driver =>
            {
                int tryCount = 0;

                while (tryCount < 3)
                {

                    var list = desktopSession.WindowHandles;
                    string xpath_LeftClickPanePnlKactuS_52_58 = "//*[contains(@Name, 'PnlKactuS')]";
                    Thread.Sleep(1000);
                    try
                    {
                        versionButtom = desktopSession.FindElementByXPath(xpath_LeftClickPanePnlKactuS_52_58);
                        versionButtom.ClearCache();
                        versionButtom.DisableCache();
                        //System.Diagnostics.Debugger.Launch();
                        versionButtom.Click();
                        Thread.Sleep(1000);
                        newFunctions_4.ScreenshotNuevo("Versión Programa", true, file);
                        rootSession = RootSession();
                        rootSession = ReloadQbeSession(rootSession, "TFrmGnAcerc");
                        if (rootSession != null)
                        {
                            versionButtom = rootSession.FindElementByXPath("//*[contains(@Name, 'Cerrar')]");
                            Thread.Sleep(2000);
                            versionButtom.Click();
                        }
                        break;
                    }
                    catch (Exception)
                    {
                        try
                        {
                            xpath_LeftClickPanePnlKactuS_52_58 = "//*[contains(@ClassName, 'TPanel')]";
                            Thread.Sleep(1000);
                            versionButtom = desktopSession.FindElementByXPath(xpath_LeftClickPanePnlKactuS_52_58);
                            versionButtom.ClearCache();
                            versionButtom.DisableCache();

                            versionButtom.Click();
                            rootSession = RootSession();
                            rootSession = ReloadQbeSession(rootSession, "TFrmGnAcerc");
                            Thread.Sleep(1000);
                            newFunctions_4.ScreenshotNuevo("Versión Programa", true, file);
                            //System.Diagnostics.Debugger.Launch();
                            if (rootSession != null)
                            {
                                versionButtom = rootSession.FindElementByXPath("//*[contains(@Name, 'Cerrar')]");
                                Thread.Sleep(2000);
                                versionButtom.Click();
                            }

                            break;
                        }
                        catch (Exception)
                        {
                            try
                            {
                                var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
                                desktopSession.Mouse.MouseMove(parentElement.Coordinates, 1764, 201);
                                desktopSession.Mouse.Click(null);
                                rootSession = RootSession();
                                rootSession = ReloadQbeSession(rootSession, "TFrmGnAcerc");
                                Thread.Sleep(1000);
                                newFunctions_4.ScreenshotNuevo("Versión Programa", true, file);
                                //System.Diagnostics.Debugger.Launch();
                                if (rootSession != null)
                                {
                                    versionButtom = rootSession.FindElementByXPath("//*[contains(@Name, 'Cerrar')]");
                                    Thread.Sleep(2000);
                                    versionButtom.Click();
                                }
                                break;
                            }
                            catch (Exception)
                            {
                                tryCount++;
                                continue;
                            }
                        }

                    }
                    break;
                }

                return versionButtom != null;
            });
        }
        public static void CloseVersion(WindowsDriver<WindowsElement> desktopSession)
        {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));

            WindowsElement versionButtom = null;
            bool brk = false;
            wait.Until(driver =>
            {
                string xpath_LeftClickPanePnlKactuS_52_58 = "";
                Thread.Sleep(1000);
                try
                {
                    //System.Diagnostics.Debugger.Launch();
                    desktopSession.FindElementByAccessibilityId("Close").Click();

                }
                catch (Exception)
                {
                    try
                    {

                        versionButtom = desktopSession.FindElement(By.Name(desktopSession.Title));
                        desktopSession.Mouse.MouseMove(versionButtom.Coordinates, 1120, 300);
                        desktopSession.Mouse.Click(null);
                        brk = true;

                    }
                    catch (Exception)
                    {


                    }
                }

                return versionButtom != null;
            });

        }


        public static List<string> validarCodPrograma(WindowsDriver<WindowsElement> desktopSession, string programa)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(rootSession)
            {
                Timeout = TimeSpan.FromSeconds(60),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));

            var ElementList = desktopSession.FindElementsByClassName("TPanel");
            var obtenerLetrasProg = programa.Substring(0, 3).ToLower();

            foreach (var elem in ElementList)
            {
                if (elem.Text.Length > 0)
                {
                    if (elem.Text.Length > 3)
                    {
                        var obtenerLetrasTpanel = elem.Text.Substring(0, 3).ToLower();
                        if (obtenerLetrasProg == obtenerLetrasTpanel)
                        {
                            if (elem.Text == programa)
                            {
                                Debug.WriteLine("Codigo del programa coincide");
                            }
                            else
                            {
                                errorMessages.Add($"Codigo de programa es diferente al esperado. Esperado: {programa}, Encontrado: {elem.Text}");
                            }
                            break;
                        }
                    }

                }
            }
            return errorMessages;
        }


        public static List<string> validarDescripPrograma(WindowsDriver<WindowsElement> desktopSession, string descripcionDada)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(60),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));

            var ElementList = desktopSession.FindElementsByClassName("Internet Explorer_Server");

            //Buscamos el tipo de servidor
            string TextoServidor = ElementList[0].Text.Remove(0, 6);

            //Encontrar posicion inicio y final
            int posicionServer0 = TextoServidor.IndexOf("/") + 1;
            int posicionServer1 = TextoServidor.IndexOf(":");

            ////Obtener nombre Server
            string descripcionServerObte = TextoServidor.Substring(posicionServer0, posicionServer1 - posicionServer0);

            if (descripcionServerObte == "dwtfskscm")
            {
                string Texto = null;
                string navigator = AbrirPrograma.SelectNavigator();
                if (navigator == "Edge")
                {
                    Texto = ElementList[0].Text.Remove(0, 72);
                }
                else
                {
                    Texto = ElementList[0].Text.Remove(0, 78);
                }

                //Encontrar posicion inicio
                int posicion0 = Texto.IndexOf("&") + 1;
                int posicion1 = Texto.IndexOf("&.");

                ////Obtener nombre
                string descripcionObte = Texto.Substring(posicion0, posicion1 - posicion0);

                //Limpiando la descripcion del programa obtenido
                string descripcionObteMostrar = descripcionObte.Replace("%20", " ");
                descripcionObte = descripcionObte.Replace("%20", "");
                descripcionObte = descripcionObte.ToLower();

                //Limpiando la descripcion del programa dado
                string descripcionDadaMostrar = descripcionDada;
                descripcionDada = descripcionDada.Replace(" ", "");
                descripcionDada = descripcionDada.ToLower();

                //if (descripcionDada != descripcionObte)
                //{
                //    errorMessages.Add($"Descripción del programa es diferente al esperado. Esperado: {descripcionDadaMostrar}, Encontrado: {descripcionObteMostrar}");
                //}
            }
            else
            {
                string Texto = ElementList[0].Text.Remove(0, 54);

                //Encontrar posicion inicio
                int posicion0 = Texto.IndexOf("&") + 1;
                int posicion1 = Texto.IndexOf("&.");

                ////Obtener nombre
                string descripcionObte = Texto.Substring(posicion0, posicion1 - posicion0);

                //Limpiando la descripcion del programa obtenido
                string descripcionObteMostrar = descripcionObte.Replace("%20", " ");
                descripcionObte = descripcionObte.Replace("%20", "");
                descripcionObte = descripcionObte.ToLower();

                //Limpiando la descripcion del programa dado
                string descripcionDadaMostrar = descripcionDada;
                descripcionDada = descripcionDada.Replace(" ", "");
                descripcionDada = descripcionDada.ToLower();

                //if (descripcionDada != descripcionObte)
                //{
                //    errorMessages.Add($"Descripción del programa es diferente al esperado. Esperado: {descripcionDadaMostrar}, Encontrado: {descripcionObteMostrar}");
                //}
            }
            return errorMessages;
        }

        public static List<string> QBEQry(WindowsDriver<WindowsElement> desktopSession, string bandera, string QbeFilter, string file,string campo=null)
        {
            List<string> errorMessages = new List<string>();
            if(bandera=="0")
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
                    List<Point> coordenadas = coordinatesFinder(desktopSession, 244, 202, 32);

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
                            }
                            else
                            {
                                elements[0].SendKeys(QbeFilter);
                            }
                        }
                        elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                        rootSession.Mouse.MouseMove(elements[8].Coordinates, 20,20);
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

        public static List<string> QBEQry2Campos(WindowsDriver<WindowsElement> desktopSession, string bandera, string QbeFilter, string QbeFilter2, string file, string campo1 = null,string campo2=null)
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
                    List<Point> coordenadas = coordinatesFinder(desktopSession, 244, 202, 32);

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
                            
                            elements[0].Click();
                            int mi = Convert.ToInt32(campo1);
                            for (int i = 1; i <= mi; i++)
                            {
                                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                            }
                            Thread.Sleep(500);
                            rootSession.Keyboard.SendKeys(QbeFilter);
                            Thread.Sleep(500);
                            newFunctions_4.ScreenshotNuevo("Identificacion QBE", true, file);
                            Thread.Sleep(500);
                            
                            int m = Convert.ToInt32(campo2);
                                for (int i = 1; i <= m; i++)
                                {
                                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                                }
                            Thread.Sleep(500);
                            rootSession.Keyboard.SendKeys(QbeFilter2);
                            Thread.Sleep(500);
                            newFunctions_4.ScreenshotNuevo("Fechas QBE", true, file);

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

            return errorMessages;
        }
        public static List<Point> coordinatesFinder(WindowsDriver<WindowsElement> session, int R, int G, int B)
        {
            Screenshot image = ((ITakesScreenshot)session).GetScreenshot();
            string name = FuncionesVitales.Hora();
            string path = "C:\\imagenesTest\\" + name + ".Png";
            Image imgSource;
            try
            {
                image.SaveAsFile(string.Format("C:\\imagenesTest\\{0}.Png", name), ScreenshotImageFormat.Png);
                Thread.Sleep(1000);
                imgSource = Image.FromFile(path);
            }
            catch (Exception e)
            {
                image.SaveAsFile(string.Format("C:\\imagenesTest\\{0}.Png", name), ScreenshotImageFormat.Png);
                Thread.Sleep(1000);
                imgSource = Image.FromFile(path);
            }

            Point puntos = new Point();
            int lastx = 0;
            int lasty = 0;
            List<Point> coordenadas = new List<Point>();
            Bitmap bmpSource = new Bitmap(imgSource);
            Bitmap bmpDest = new Bitmap(bmpSource.Width, bmpSource.Height);
            int romper = 0;
            for (int y = 0; y < bmpSource.Height; y++)
            {
                for (int x = 0; x < bmpSource.Width; x++)
                {
                    Color clrPixel = bmpSource.GetPixel(x, y);
                    if (clrPixel.R == R && clrPixel.G == G && clrPixel.B == B)
                    {
                        puntos.X = x;
                        puntos.Y = y;
                        if (x > lastx + 10 || y > lasty + 10)
                        {
                            coordenadas.Add(puntos);
                            lastx = x;
                            lasty = y;
                            romper++;
                            break;
                        }
                    }
                }
                if (romper != 0)
                {
                    break;
                }
            }
            return coordenadas;
        }


        public static List<string> ReportePreliminar(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1, string nomPrograma = null)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            if(bandera == "0" || bandera == "1")
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
                    List<Point> coordenadas = coordinatesFinder(desktopSession, 227, 117, 0);
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
                            errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1, nomPrograma);
                            if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                            
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

        public static void RecursivelyClickAllElements(WindowsDriver<WindowsElement> desktopSession, string window, WindowsElement e, string xPath)
        {

            xPath += ConstructXPathFromElement(e);

            System.Diagnostics.Debug.WriteLine($"Control: {e.GetAttribute("AutomationId")}; Displayed: {e.Displayed}");


            if (e.Displayed)
                e.Click();


            string currentWindow = desktopSession.CurrentWindowHandle;
            if (currentWindow != window)
            {

            }


            var childElements = desktopSession.FindElementsByXPath($"{xPath}*");


            if (childElements == null)
                return;
            else if (childElements.Count == 0)
                return;


            foreach (var element in childElements)
            {
                RecursivelyClickAllElements(desktopSession, currentWindow, element, xPath);
            }
        }

        public static string ConstructXPathFromElement(WindowsElement e)
        {
            string className = e.GetAttribute("LocalizedControlType");
            className = (char.ToUpper(className[0]) + className.Substring(1));
            string automationId = e.GetAttribute("AutomationId");
            string name = e.GetAttribute("Name");
            string xPath = $"*[@AutomationId=\"{automationId}\"][@Name=\"{name}\"]/";

            return xPath;
        
        }





        public static CodeCompileUnit BuildHelloWorldGraph()
        {
            // Create a new CodeCompileUnit to contain
            // the program graph.
            CodeCompileUnit compileUnit = new CodeCompileUnit();

            // Declare a new namespace called Samples.
            CodeNamespace samples = new CodeNamespace("Samples");
            // Add the new namespace to the compile unit.
            compileUnit.Namespaces.Add(samples);

            // Add the new namespace import for the System namespace.
            samples.Imports.Add(new CodeNamespaceImport("System"));

            // Declare a new type called Class1.
            CodeTypeDeclaration class1 = new CodeTypeDeclaration("Class1");
            // Add the new type to the namespace type collection.
            samples.Types.Add(class1);

            // Declare a new code entry point method.
            CodeEntryPointMethod start = new CodeEntryPointMethod();

            // Create a type reference for the System.Console class.
            CodeTypeReferenceExpression csSystemConsoleType = new CodeTypeReferenceExpression("System.Console");

            // Build a Console.WriteLine statement.
            CodeMethodInvokeExpression cs1 = new CodeMethodInvokeExpression(
                csSystemConsoleType, "WriteLine",
                new CodePrimitiveExpression("Hello World!"));

            // Add the WriteLine call to the statement collection.
            start.Statements.Add(cs1);
            // Build another Console.WriteLine statement.
            CodeMethodInvokeExpression cs2 = new CodeMethodInvokeExpression(
                csSystemConsoleType, "WriteLine",
                new CodePrimitiveExpression("Press the Enter key to continue."));

            // Add the WriteLine call to the statement collection.
            start.Statements.Add(cs2);

            // Build a call to System.Console.ReadLine.
            CodeMethodInvokeExpression csReadLine = new CodeMethodInvokeExpression(
                csSystemConsoleType, "ReadLine");

            // Add the ReadLine statement.
            start.Statements.Add(csReadLine);

            // Add the code entry point method to
            // the Members collection of the type.
            class1.Members.Add(start);



            TestMethodAttribute test = new TestMethodAttribute();

            //test.Execute();


            return compileUnit;
        }
        
        public static void GenerateCode(CodeDomProvider provider,
            CodeCompileUnit compileunit)
        {
            // Build the source file name with the appropriate
            // language extension.
            String sourceFile;
            if (provider.FileExtension[0] == '.')
            {
                sourceFile = "TestGraph" + provider.FileExtension;
            }
            else
            {
                sourceFile = "TestGraph." + provider.FileExtension;
            }

            // Create an IndentedTextWriter, constructed with
            // a StreamWriter to the source file.
            IndentedTextWriter tw = new IndentedTextWriter(new StreamWriter(sourceFile, false), "    ");
            // Generate source code using the code generator.
            provider.GenerateCodeFromCompileUnit(compileunit, tw, new CodeGeneratorOptions());
            // Close the output file.
            tw.Close();
        }

        public static CompilerResults CompileCode(CodeDomProvider provider,
                                                  String sourceFile,
                                                  String exeFile)
        {
            // Configure a CompilerParameters that links System.dll
            // and produces the specified executable file.
            String[] referenceAssemblies = { "System.dll" };
            CompilerParameters cp = new CompilerParameters(referenceAssemblies,
                                                           exeFile, false);
            // Generate an executable rather than a DLL file.
            cp.GenerateExecutable = true;

            // Invoke compilation.
            CompilerResults cr = provider.CompileAssemblyFromFile(cp, sourceFile);
            // Return the results of compilation.
            return cr;
        }

        public static void CerrarVentanaNotifiacion(WindowsDriver<WindowsElement> desktopSession)
        {
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(rootSession)
            {
                Timeout = TimeSpan.FromSeconds(60),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            
            var ElementList = desktopSession.FindElementsByClassName("BubbleFrameView");
            //var ElementList1 = desktopSession.FindElementsByClassName("NonClientView");

            if (ElementList.Count > 0)
            {
                for (int i = 0; i < 1; i++)
                {
                    ElementList = desktopSession.FindElementsByClassName("BubbleFrameView");
                    ElementList[0].FindElementByName("Cerrar").Click();
                    Thread.Sleep(3000);
                }
                Thread.Sleep(1000);
            }
        }

    }
}
