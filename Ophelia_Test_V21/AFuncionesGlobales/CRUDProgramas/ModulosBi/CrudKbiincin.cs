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
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBi
{
    class CrudKbiincin : FuncionesVitales
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


        static public void CrudPrincipal(WindowsDriver<WindowsElement> desktopSession, List<string> crudPrincipal)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");
            var btnTDVI1 = desktopSession.FindElementsByClassName("TGridPanel");

            //string nro = "100";
            //string nro1 = "76";
            desktopSession.Mouse.MouseMove(btnTDVI1[0].Coordinates, 261, 21);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(500);
            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 54, 45);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
            Thread.Sleep(1000);        
            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            Thread.Sleep(1000);
            desktopSession.Keyboard.SendKeys(crudPrincipal[2]);
            Thread.Sleep(1000);
            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
            Thread.Sleep(1000);
        }

        

        static public void CrudDet2(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDet2)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            switch (bandera)
            {
                case 1:

                    // Datos para add registro
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 169, 354);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                    Thread.Sleep(500);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(500);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    desktopSession.Keyboard.SendKeys(crudDet2[0]);
                    Thread.Sleep(500);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    desktopSession.Keyboard.SendKeys(crudDet2[1]);
                    Thread.Sleep(500);

                    break;

                case 2:

                    // Datos para update registro
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 193, 341);
                    Thread.Sleep(500);
                    desktopSession.Mouse.DoubleClick(null);
                    Thread.Sleep(500);
                    desktopSession.Keyboard.SendKeys(crudDet2[2]);
                    Thread.Sleep(500);

                    break;

            }         
            
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


        //preliminar 1
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
                        Thread.Sleep(400);
                        WindowsDriver<WindowsElement> RootSession = null;
                        RootSession = PruebaCRUD.RootSession();
                        // Flecha hacia bajo en la ventana previa del preli.
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        newFunctions_4.ScreenshotNuevo("Preliminar 1", true, file);
                        Thread.Sleep(400);
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        Thread.Sleep(4500);

                        if (bandera == "0")
                        {
                            errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1, nomPrograma);
                            //if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
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
                        RootSession = ReloadSession(RootSession, "TFrmReporte");


                        // Flecha hacia bajo en la ventana previa del preli.
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                        Thread.Sleep(400);

                        // Click en primera opcion preliminar 2
                        var btnTDVI2 = RootSession.FindElementsByClassName("TPanel");

                        RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 11, 28);
                        RootSession.Mouse.Click(null);
                        Thread.Sleep(2000);

                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(400);

                        // in fechas para poder hacer report preliminar
                        RootSession.Keyboard.SendKeys("1/10/2022");
                        Thread.Sleep(1000);
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(400);
                        RootSession.Keyboard.SendKeys("31/10/2022");
                        Thread.Sleep(1000);

                        // SS Preliminar 2.
                        newFunctions_4.ScreenshotNuevo("Preliminar 2", true, file);
                        Thread.Sleep(400);
                        
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(400);

                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(400);

                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        Thread.Sleep(2500);

                        if (bandera == "0")
                        {
                            errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1, nomPrograma);
                            //if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
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

        //preliminar 3
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
                        RootSession = ReloadSession(RootSession, "TFrmReporte");


                        // Flecha hacia bajo en la ventana previa del preli.
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);


                        // Click en primera opcion preliminar 2
                        var btnTDVI2 = RootSession.FindElementsByClassName("TPanel");

                        RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 11, 28);
                        RootSession.Mouse.Click(null);
                        Thread.Sleep(2000);

                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                        Thread.Sleep(400);
                        

                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(400);

                        // in fechas para poder hacer report preliminar
                        RootSession.Keyboard.SendKeys("1/10/2022");
                        Thread.Sleep(1000);
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(400);
                        RootSession.Keyboard.SendKeys("31/10/2022");
                        Thread.Sleep(1000);

                        // SS Preliminar 2.
                        newFunctions_4.ScreenshotNuevo("Preliminar 3", true, file);
                        Thread.Sleep(400);

                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(400);

                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(400);

                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        Thread.Sleep(2500);

                        if (bandera == "0")
                        {
                            errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1, nomPrograma);
                            //if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
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

        //preliminar 4
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
                        RootSession = ReloadSession(RootSession, "TFrmReporte");


                        // Flecha hacia bajo en la ventana previa del preli.
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);


                        // Click en primera opcion preliminar 2
                        var btnTDVI2 = RootSession.FindElementsByClassName("TPanel");

                        RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 11, 28);
                        RootSession.Mouse.Click(null);
                        Thread.Sleep(2000);

                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                        Thread.Sleep(400);
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                        Thread.Sleep(400);


                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(400);

                        // in fechas para poder hacer report preliminar
                        RootSession.Keyboard.SendKeys("1/10/2022");
                        Thread.Sleep(1000);
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(400);
                        RootSession.Keyboard.SendKeys("31/10/2022");
                        Thread.Sleep(1000);

                        // SS Preliminar 2.
                        newFunctions_4.ScreenshotNuevo("Preliminar 4", true, file);
                        Thread.Sleep(400);

                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(400);

                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(400);

                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        Thread.Sleep(2500);

                        if (bandera == "0")
                        {
                            errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1, nomPrograma);
                           // if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
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

        //preliminar 5
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
                        RootSession = ReloadSession(RootSession, "TFrmReporte");


                        // Flecha hacia bajo en la ventana previa del preli.
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);


                        // Click en primera opcion preliminar 2
                        var btnTDVI2 = RootSession.FindElementsByClassName("TPanel");

                        RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 11, 28);
                        RootSession.Mouse.Click(null);
                        Thread.Sleep(2000);

                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                        Thread.Sleep(400);
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                        Thread.Sleep(400);
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                        Thread.Sleep(400);

                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(400);

                        // in fechas para poder hacer report preliminar
                        RootSession.Keyboard.SendKeys("1/10/2022");
                        Thread.Sleep(1000);
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(400);
                        RootSession.Keyboard.SendKeys("31/10/2022");
                        Thread.Sleep(1000);

                        // SS Preliminar 2.
                        newFunctions_4.ScreenshotNuevo("Preliminar 5", true, file);
                        Thread.Sleep(400);

                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(400);

                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(400);

                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        Thread.Sleep(2500);

                        if (bandera == "0")
                        {
                            errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1, nomPrograma);
                           // if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
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

        //preliminar 6
        public static List<string> ReportePreliminar6(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1, string nomPrograma = null)
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
                        RootSession = ReloadSession(RootSession, "TFrmReporte");


                        // Flecha hacia bajo en la ventana previa del preli.
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);


                        // Click en primera opcion preliminar 2
                        var btnTDVI2 = RootSession.FindElementsByClassName("TPanel");

                        RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 11, 28);
                        RootSession.Mouse.Click(null);
                        Thread.Sleep(2000);


                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                        Thread.Sleep(400);
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                        Thread.Sleep(400);
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                        Thread.Sleep(400);
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                        Thread.Sleep(400);


                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(400);

                        // in fechas para poder hacer report preliminar
                        RootSession.Keyboard.SendKeys("1/10/2022");
                        Thread.Sleep(1000);
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(400);
                        RootSession.Keyboard.SendKeys("31/10/2022");
                        Thread.Sleep(1000);

                        // SS Preliminar 2.
                        newFunctions_4.ScreenshotNuevo("Preliminar 6", true, file);
                        Thread.Sleep(400);

                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(400);

                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(400);

                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        Thread.Sleep(2500);

                        if (bandera == "0")
                        {
                            errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1, nomPrograma);
                            //if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
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
