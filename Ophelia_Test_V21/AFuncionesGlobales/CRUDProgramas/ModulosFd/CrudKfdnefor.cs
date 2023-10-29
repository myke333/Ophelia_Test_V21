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
using System.Drawing.Imaging;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosFd
{
    class CrudKfdnefor : FuncionesVitales
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
                    rootSession = ReloadSession(rootSession, "TFrmSeleccion");

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

                    var btn = rootSession.FindElementsByClassName("TButton");
                    ////Debugger.Launch();
                    foreach (var boton in btn)
                    {
                        if (boton.Text == "Aceptar")
                        {
                            boton.Click();
                                rootSession = RootSession();
                                rootSession = ReloadSession(rootSession, "TFrmFdNxJef");
                           // //Debugger.Launch();
                            //bool ventana1 = coordinatesFinder(desktopSession, 27, 101, 206, screenWidth / 2, 347, 1);
                            if (rootSession != null)
                            {
                               // //Debugger.Launch();

                                var btn1 = rootSession.FindElementsByClassName("TBitBtn");

                                  QBEQry(desktopSession, "0", crudVariables, file);
                                   foreach (var boton1 in btn1)
                                    {
                                        if (boton1.Text == "Aceptar")
                                        {
                                            boton1.Click();

                                         }
                                     }
                                 }                                      
                            }
                    }
                    bool ventana = coordinatesFinder(desktopSession, 27, 101, 206, screenWidth / 2, 347, 1);

                    if (ventana == true)
                    {
                        screenWidth = Screen.PrimaryScreen.Bounds.Width;
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        Thread.Sleep(2000);

                        //bool ventana = coordinatesFinder(desktopSession, 27, 101, 206, screenWidth / 2, 347, 1);
                        rootSession = RootSession();
                        rootSession = ReloadSession(rootSession, "IEFrame");
                        var allFrame = rootSession.FindElementsByClassName("IEFrame");
                        new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset(allFrame[0].Size.Width / 2, (allFrame[0].Size.Height / 2) + 30).DoubleClick().Click().Perform();


                    }

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

        public static List<Point> coordinatesFinderAux(WindowsDriver<WindowsElement> session, int R, int G, int B, string name)
        {

            string path = "C:\\imagenesTest\\" + name + ".Png";
            Image imgSource;
            try
            {
                int screenWidth = Screen.PrimaryScreen.Bounds.Width;
                int screenHeight = Screen.PrimaryScreen.Bounds.Height;




                Bitmap captureBitmap = new Bitmap(screenWidth, screenHeight, PixelFormat.Format32bppArgb);

                Rectangle captureRectangle = Screen.AllScreens[0].Bounds;

                Graphics captureGraphics = Graphics.FromImage(captureBitmap);

                captureGraphics.CopyFromScreen(captureRectangle.Left, captureRectangle.Top, 0, 0, captureRectangle.Size);

                captureBitmap = (new Bitmap(captureBitmap, new Size(1600, 900)));

                captureBitmap.Save(string.Format(path, name), ImageFormat.Bmp);
                imgSource = Image.FromFile(path);
            }
            catch (Exception e)
            {

                imgSource = Image.FromFile(path);
            }

            Point puntos = new Point();
            int lastx = 0;
            int lasty = 0;
            List<Point> coordenadas = new List<Point>();
            Bitmap bmpSource = new Bitmap(imgSource);
            Bitmap bmpDest = new Bitmap(bmpSource.Width, bmpSource.Height);
            coordenadas = PixelSeeker(bmpSource, R, G, B, new Point(20, 20));

            //if (coordenadas.Count == 0) coordenadas = PixelSeeker(bmpSource, R, G, B, new Point(20, 20));

            imgSource.Dispose();

            ////Debugger.Launch();
            return coordenadas;

        }

        public static List<Point> PixelSeeker(Bitmap image, int R, int G, int B, Point offset)
        {

            Point punto = new Point();
            int lastx = 0;
            int lasty = 0;
            List<Point> coordinates = new List<Point>();

            for (int y = 0; y < image.Height; y++)
            {
                for (int x = 0; x < image.Width; x++)
                {
                    Color clrPixel = image.GetPixel(x, y);
                    if (clrPixel.R == R && clrPixel.G == G && clrPixel.B == B)
                    {
                        punto.X = x;
                        punto.Y = y;
                        if (x > lastx + offset.X || y > lasty + offset.Y)
                        {
                            coordinates.Add(punto);
                            lastx = x;
                            lasty = y;
                        }
                    }
                }
            }


            image.Dispose();
            return coordinates;

        }

        public static List<string> QBEQry(WindowsDriver<WindowsElement> desktopSession, string bandera, List<string> crudVariables, string file)
        {
            string name = "PRUEBA";
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
                    List<Point> coordenadas = coordinatesFinderAux(desktopSession, 244, 202, 32,name);

                    int coordX = coordenadas[1].X;
                    int coordY = coordenadas[1].Y;
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
                        if (!string.IsNullOrEmpty(crudVariables[10]))
                        {
                            elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/Pane[contains(@ClassName, 'TStringGrid')]/*"));
                            elements[0].SendKeys(crudVariables[10]);
                        }
                        elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                        rootSession.Mouse.MouseMove(elements[8].Coordinates, 20, 20);
                        rootSession.Mouse.Click(null);
                        Thread.Sleep(500);

                        if (string.IsNullOrEmpty(crudVariables[10]))
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

       
        private static bool coordinatesFinder(WindowsDriver<WindowsElement> session, int R, int G, int B, int fil, int col, int tipo)
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

            Bitmap bmpSource = new Bitmap(imgSource);
            Bitmap bmpDest = new Bitmap(bmpSource.Width, bmpSource.Height);
            bool val = false;

            //Color clrPixel = bmpSource.GetPixel(fil, col);

            if (tipo == 2)
            {
                Color clrPixel = bmpSource.GetPixel(fil, col);
                if (clrPixel.R == R && clrPixel.G == G && clrPixel.B == B)
                {
                    val = true;
                }
            }
            else if (tipo == 1)
            {
                int cont = 0;
                for (int x = 0; x < bmpSource.Width; x++)
                {
                    for (int y = 0; y < bmpSource.Height; y++)
                    {
                        Color clrPixel = bmpSource.GetPixel(x, y);
                        if (cont == 0)
                        {
                            if (clrPixel.R == 203 && clrPixel.G == 219 && clrPixel.B == 234)
                            {
                                cont = 1;
                                val = false;
                                break;
                            }
                        }
                    }
                    if (cont != 0) { break; }
                }
                if (cont == 0) { val = true; }
            }

            return val;
        }

        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVariables, String file)
        {
            List<string> errorMessages = new List<string>();
            int screenWidth = Screen.PrimaryScreen.Bounds.Width;
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            List<int> lupaOff = new List<int>() { 6 };
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementBox = desktopSession.FindElementsByClassName("TDBCmbCodigoBox");
            

            if (tipo == 1)
            {

                {
                    newFunctions_4.LupaDinamicaDiscriminatoria(desktopSession, crudVariables,lupaOff);

                    ElementList[2].SendKeys(crudVariables[6]);                    
                    ElementBox[0].SendKeys(crudVariables[7]);
                    ElementList[17].SendKeys(crudVariables[8]);
                }

            }

            else
            {
                ElementList[17].Clear();
                ElementList[17].SendKeys(crudVariables[9]);
            }
        }
    }
}
