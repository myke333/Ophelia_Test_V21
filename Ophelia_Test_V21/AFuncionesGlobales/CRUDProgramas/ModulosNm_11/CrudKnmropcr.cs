using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Drawing;
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
using System.Drawing.Imaging;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;




namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_11
{
    class CrudKnmropcr : FuncionesVitales
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

        static public void SeleccionDatos(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal)
        {
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");

            if (bandera == 0)
            {


                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 95, 31);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                Thread.Sleep(1000);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 205, 31);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                Thread.Sleep(1000);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 380, 31);
                desktopSession.Mouse.DoubleClick(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                desktopSession.Mouse.DoubleClick(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                desktopSession.Mouse.DoubleClick(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                desktopSession.Keyboard.SendKeys(crudPrincipal[2]);
                Thread.Sleep(1000);
            }
        }

        //static public void Click(WindowsDriver<WindowsElement> desktopSession, string file)
        //{
        //    var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");

        //    desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 415, 335);
        //    desktopSession.Mouse.Click(null);
        //    Thread.Sleep(1500);
        //    newFunctions_4.ScreenshotNuevo("Click QBE Det 1", true, file);
        //    //desktopSession.Mouse.Click(null);
        //    //Thread.Sleep(1500);

        //    WindowsDriver<WindowsElement> RootSession = null;
        //    RootSession = PruebaCRUD.RootSession();
        //    RootSession = ReloadSession(RootSession, "TKQBE");

        //    var btnTDVI = RootSession.FindElementsByClassName("TTabSheet");

        //    RootSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 600, 16);
        //    RootSession.Mouse.Click(null);
        //    Thread.Sleep(2000);

        //    WindowsDriver<WindowsElement> RootSession2 = null;
        //    RootSession2 = PruebaCRUD.RootSession();
        //    RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
        //    Thread.Sleep(2000);
        //    newFunctions_4.ScreenshotNuevo("Registros QBE", true, file);

        //    WindowsDriver<WindowsElement> RootSession3 = null;
        //    RootSession3 = PruebaCRUD.RootSession();
        //    RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
        //    Thread.Sleep(2000);
        //}


        //public static void ClickButtonAceptarPreliminar(WindowsDriver<WindowsElement> desktopSession, string file)
        //{
        //    var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");

        //    newFunctions_4.ScreenshotNuevo("Se da Click en el boton Preliminar y luego en Aceptar", true, file);
        //    WindowsDriver<WindowsElement> RootSession6 = null;
        //    RootSession6 = PruebaCRUD.RootSession();
        //    RootSession6.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
        //}
        //

        
        public static List<string> ReportePreliminar1(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1, string nomPrograma = null)
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
                    List<Point> coordenadas = coordinatesFinder(desktopSession, 227, 117, 0);
                    int coordX = coordenadas[0].X;
                    int coordY = coordenadas[0].Y;
                    var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
                    Thread.Sleep(1000);


                    try
                    {
                        desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordX, coordY);
                        desktopSession.Mouse.Click(null);
                        Thread.Sleep(1800);

                        //newFunctions_4.ScreenshotNuevo("Preliminar 1 - Coste rela del gasto/ Normal", true, file);

                        WindowsDriver<WindowsElement> RootSession = null;
                        RootSession = PruebaCRUD.RootSession();
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        Thread.Sleep(1000);
                        
                        Thread.Sleep(4500);

                        if (bandera == "0")
                        {
                            errors = newFunctions_5.GenerarReportes("Preliminar 10 ", file, pdf1, nomPrograma);
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

        //preliminar 2
        public static List<string> ReportePreliminar2(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1, string nomPrograma = null)
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
                    List<Point> coordenadas = coordinatesFinder(desktopSession, 227, 117, 0);
                    int coordX = coordenadas[0].X;
                    int coordY = coordenadas[0].Y;
                    var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
                    Thread.Sleep(1000);


                    try
                    {
                        desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordX, coordY);
                        desktopSession.Mouse.Click(null);
                        Thread.Sleep(400);

                        WindowsDriver<WindowsElement> RootSession = null;
                        RootSession = PruebaCRUD.RootSession();
                                               
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(500);
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(500);
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(500);

                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                        Thread.Sleep(500);
                        //newFunctions_4.ScreenshotNuevo("Preliminar 2 - Costo real de gasto / centro de costo ", true, file);
                        //Thread.Sleep(400);
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        Thread.Sleep(4500);

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

        public static List<string> ReportePreliminar3(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1, string nomPrograma = null)
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
                    List<Point> coordenadas = coordinatesFinder(desktopSession, 227, 117, 0);
                    int coordX = coordenadas[0].X;
                    int coordY = coordenadas[0].Y;
                    var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
                    Thread.Sleep(1000);


                    try
                    {
                        desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordX, coordY);
                        desktopSession.Mouse.Click(null);
                        Thread.Sleep(400);
                        WindowsDriver<WindowsElement> RootSession = null;
                        RootSession = PruebaCRUD.RootSession();
                       
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                        Thread.Sleep(500);

                        //newFunctions_4.ScreenshotNuevo("Preliminar 3 / Orden de Pago ", true, file);
                        Thread.Sleep(400);
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        Thread.Sleep(500);
                        
                        
                        ////Cerrar esta ventana
                        //WindowsDriver<WindowsElement> RootSession2 = null;
                        //RootSession2 = PruebaCRUD.RootSession();
                        //RootSession2 = ReloadSession(RootSession2, "TFrmValNive");
                        //Thread.Sleep(400);

                        //RootSession2.FindElementByName("Aceptar").Click();
                        //Thread.Sleep(1000);0


                        if (bandera == "0")
                        {
                            errors = newFunctions_5.GenerarReportes("Preliminar3", file, pdf1, nomPrograma);
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


        //preliminar 2
        public static List<string> ReportePreliminar4(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1, string nomPrograma = null)
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
                    List<Point> coordenadas = coordinatesFinder(desktopSession, 227, 117, 0);
                    int coordX = coordenadas[0].X;
                    int coordY = coordenadas[0].Y;
                    var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
                    Thread.Sleep(1000);


                    try
                    {
                        desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordX, coordY);
                        desktopSession.Mouse.Click(null);
                        Thread.Sleep(400);

                        WindowsDriver<WindowsElement> RootSession = null;
                        RootSession = PruebaCRUD.RootSession();
                        
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                        Thread.Sleep(500);

                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                        Thread.Sleep(500);

                        //newFunctions_4.ScreenshotNuevo("Preliminar 4 - Resumen ordenes de pago / Normal ", true, file);
                        //Thread.Sleep(500);

                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        Thread.Sleep(500);

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


        //preliminar 2
        public static List<string> ReportePreliminar5(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1, string nomPrograma = null)
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
                    List<Point> coordenadas = coordinatesFinder(desktopSession, 227, 117, 0);
                    int coordX = coordenadas[0].X;
                    int coordY = coordenadas[0].Y;
                    var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
                    Thread.Sleep(1000);


                    try
                    {
                        desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordX, coordY);
                        desktopSession.Mouse.Click(null);
                        Thread.Sleep(400);

                        WindowsDriver<WindowsElement> RootSession = null;
                        RootSession = PruebaCRUD.RootSession();

                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                        Thread.Sleep(500);
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                        Thread.Sleep(500);

                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(500);
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(500);
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(500);
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                        Thread.Sleep(500);

                        //newFunctions_4.ScreenshotNuevo("Preliminar 5 - Resumen ordenes de pago / centro de costo ", true, file);
                        //Thread.Sleep(400);
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        Thread.Sleep(4500);

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



    }
}