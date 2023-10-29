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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSo
{
    class CrudKsoinfac : FuncionesVitales
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

        private static List<Point> CoordinatesFinder(WindowsDriver<WindowsElement> session, int R, int G, int B)
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
            int Romper = 0;

            for (int x = 0; x < bmpSource.Width; x++)
            {
                for (int y = 0; y < bmpSource.Height; y++)
                {
                    Color clrPixel = bmpSource.GetPixel(x, y);
                    if (clrPixel.R == R && clrPixel.G == G && clrPixel.B == B)
                    {
                        puntos.X = x;
                        puntos.Y = y;
                        if (x > lastx + 20 || y > lasty + 20)
                        {
                            coordenadas.Add(puntos);
                            lastx = x;
                            lasty = y;
                            Romper++;
                            break;
                        }
                    }
                }
                if (Romper != 0)
                {
                    break;
                }
            }
            ////Debugger.Launch();
            return coordenadas;
        }
        public static List<string> ReporteDinamico(WindowsDriver<WindowsElement> desktopSession, string file, string ruta, string nomPrograma = null)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();

            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));

            try
            {
                wait.Until(driver =>
                {
                    WindowsDriver<WindowsElement> rootSession = null;

                    List<Point> coordenadas = CoordinatesFinder(desktopSession, 241, 102, 57);
                    int CoordX = coordenadas[0].X;
                    int CoordY = coordenadas[0].Y;

                    var PanelElements = desktopSession.FindElement(By.Name(desktopSession.Title));
                    bool IsDisplayedReport = false;

                    // Clic en el icono de reporte dinamico
                    desktopSession.Mouse.MouseMove(PanelElements.Coordinates, CoordX, CoordY);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(3000);

                    rootSession = RootSession();
                    rootSession = ReloadSession(rootSession, "TFrmUDRepor");
                    var elements = rootSession.FindElementsByClassName("TFrmUDRepor");

                    if (elements.Count > 0)
                    {
                        IsDisplayedReport = true;
                    }
                    else
                    {
                        errorMessages.Add($"No se puede visualizar la ventana Generar Reporte Dinámico");
                    }

                    if (IsDisplayedReport == true)
                    {
                        try
                        {

                            var ReportElements = rootSession.FindElementsByClassName("TKsListCodigo");

                            if (ReportElements.Count > 0)
                            {
                                // Agregar Variables para el reporte
                                rootSession.Mouse.MouseMove(ReportElements[1].Coordinates, 76, 7);
                                rootSession.Mouse.DoubleClick(null);
                                rootSession.Mouse.DoubleClick(null);

                                rootSession = ReloadSession(rootSession, "TFrmUDRepor");
                                // Generar reporte
                                var ElementGenerar = rootSession.FindElementByName("Generar");
                                ElementGenerar.Click();
                                Thread.Sleep(5000);
                                newFunctions_4.ScreenshotNuevo("Datos Seleccionados para Generar el Reporte Dinámico", true, file);

                                // Generar Preview
                                var ElementPreview = rootSession.FindElementByName("Preview");
                                ElementPreview.Click();
                                Thread.Sleep(3000);
                                newFunctions_4.ScreenshotNuevo("Preview del Reporte Dinámico Generado", true, file);

                                // Generar Reportes
                                errors = newFunctions_5.GenerarReportes("Dinámico", file, ruta, nomPrograma);
                                if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }

                                // Salir ventana "Generar Reporte"
                                rootSession = RootSession();
                                rootSession = ReloadSession(rootSession, "TFrmUDRepor");
                                var ElementSalir = rootSession.FindElementByName("Salir");
                                ElementSalir.Click();
                            }
                        }
                        catch
                        {
                            errorMessages.Add($"Fallo en el Proceso");
                        }
                    }
                    return rootSession != null;
                });
            }
            catch
            {

            }
            return errorMessages;
        }

        public static List<string> QBEQry(WindowsDriver<WindowsElement> desktopSession, string bandera, string QbeFilter, string file)
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
                    rootSession = ReloadSession(rootSession, "TKQBE");
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
                        if (QbeFilter!="0")
                        {
                            elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/Pane[contains(@ClassName, 'TStringGrid')]/*"));
                            elements[0].SendKeys(QbeFilter);
                        }
                        elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                        rootSession.Mouse.MouseMove(elements[8].Coordinates, 20, 20);
                        rootSession.Mouse.Click(null);
                        Thread.Sleep(500);

                        if (QbeFilter=="0")
                        {
                            rootSession = ReloadSession(rootSession, "TKQBE");
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

        public static List<string> ReportePreliminar(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1, int pos, List<string> crudVariables)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            int screenWidth = Screen.PrimaryScreen.Bounds.Width;
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

                    rootSession = RootSession();
                    rootSession = ReloadSession(rootSession, "TFrmMenu");

                    var btnRad = rootSession.FindElementsByClassName("TGroupButton");
                    int c = 0;
                    ////Debugger.Launch();
                    foreach (var btRd in btnRad)
                    {
                        if (pos == c)
                        {
                            
                            {
                                btRd.Click();
                               
                            }
                        }

                        c = c + 1;
                    }

                    var btn = rootSession.FindElementsByClassName("TBitBtn");
                    ////Debugger.Launch();
                    foreach (var boton in btn)
                    {
                        if (boton.Text == "Aceptar")
                        {
                            boton.Click();
                        }
                    }
                    //bool ventana = coordinatesFinder(desktopSession, 27, 101, 206, screenWidth / 2, 347, 1);

                    //if (ventana == true)
                    //{
                    //    screenWidth = Screen.PrimaryScreen.Bounds.Width;
                    //    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    //    Thread.Sleep(2000);

                    //    //bool ventana = coordinatesFinder(desktopSession, 27, 101, 206, screenWidth / 2, 347, 1);
                    //    rootSession = RootSession();
                    //    rootSession = ReloadSession(rootSession, "IEFrame");
                    //    var allFrame = rootSession.FindElementsByClassName("IEFrame");
                    //    new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset(allFrame[0].Size.Width / 2, (allFrame[0].Size.Height / 2) + 30).DoubleClick().Click().Perform();


                    //}

                    Thread.Sleep(2000);
                    if (bandera == "0")
                    {
                        errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
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
            return errorMessages;
        }




        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVariables, String file)
        {
            List<string> errorMessages = new List<string>();
            int screenWidth = Screen.PrimaryScreen.Bounds.Width;
            WindowsDriver<WindowsElement> rootSession = null;
            List<int> lupasOff = new List<int>() { 0 };
            rootSession = RootSession();
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementButton = desktopSession.FindElementsByClassName("TDBCmbCodigoBox");


            if (tipo == 1)
            {

                {

                   // newFunctions_4.LupaDinamicaDiscriminatoria(desktopSession, crudVariables, lupasOff);
                    ElementList[7].SendKeys(crudVariables[0]);
                    ElementList[8].SendKeys(crudVariables[1]);
                    ElementList[3].SendKeys(crudVariables[2]);


                }

            }

            else
            {

                ElementList[3].Clear();
                ElementList[3].SendKeys(crudVariables[3]);

            }
        }
       
    }
}
