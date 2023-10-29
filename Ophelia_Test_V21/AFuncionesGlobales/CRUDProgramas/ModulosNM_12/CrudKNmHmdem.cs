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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_12
{
    class CrudKNmHmdem : FuncionesVitales
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

        static public WindowsDriver<WindowsElement> ReloadSession1(WindowsDriver<WindowsElement> session, string parametro)
        {
            try
            {
                var ApplicationWindow = session.FindElementByClassName(parametro);
                var ApplicationSessionHandle = ApplicationWindow.GetAttribute("NativeWindowHandle");
                ApplicationSessionHandle = (int.Parse(ApplicationSessionHandle)).ToString("x"); //Convert to hex


                var appCapabilities = new AppiumOptions();
                appCapabilities.AddAdditionalCapability("appTopLevelWindow", ApplicationSessionHandle);
                var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);

                string ses = ApplicationSession.PageSource;
                return ApplicationSession;
            }
            catch (Exception ex) { return null; }
        }

        static public void AgregarCrud(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal, string file)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");


            if (bandera == 0)
            {

                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 135, 63);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 260, 105);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 60, 184);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[2]);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 60, 230);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 60, 277);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[4]);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 600, 184);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[5]);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 600, 230);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[6]);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 600, 277);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[7]);
                Thread.Sleep(1000);

            }
            else {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 600, 277);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[8]);
                Thread.Sleep(1000);

            }


        }

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
                        Thread.Sleep(700);
                        newFunctions_4.ScreenshotNuevo("Preliminar 1 - Preliminar", true, file);
                        WindowsDriver<WindowsElement> RootSession = null;
                        RootSession = PruebaCRUD.RootSession();
                        RootSession.Mouse.MouseMove(null, 730, 354);
                        RootSession.Mouse.Click(null);
                        Thread.Sleep(4500);
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(4500);
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





        static public void ExpExel(WindowsDriver<WindowsElement> Session, string ruta)
        {


            //Abre ventana excel
            Session = RootSession();
            Session = ReloadSession1(Session, "XLMAIN");
            if (Session != null)
            {
                int count = 0;
                while (count < 240)
                {
                    try
                    {
                        //Session = RootSession();
                        //Session = ReloadSession1(Session, "XLMAIN");
                        Thread.Sleep(500);
                        var saveExcel1 = Session.FindElementsByName("Maximizar");
                        if (saveExcel1.Count > 0)
                        {
                            break;
                        }
                    }
                    catch
                    {
                        Thread.Sleep(1000);
                        count++;
                    }
                }
                var saveExcel = Session.FindElementsByName("Maximizar");
                var cantidad = saveExcel.Count;
                if (cantidad == 2)
                {
                    saveExcel[1].Click();
                }
                else
                {
                    saveExcel[0].Click();
                }
                //Cambio Jose
                Session.FindElementByName("Pestaña Archivo").Click();
                //Fin Cambio Jose

                //Session.FindElementByName("Guardar").Click();
                Session.FindElementByName("Guardar").Click();
                Session.FindElementByName("Examinar").Click();
                Thread.Sleep(1000);
                Session.Keyboard.SendKeys(ruta);
                Session.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);
                //newFunctions_4.ScreenshotNuevo("Reporte Excel Guardado", true, file);
                LimpiarProcesos();

            }
            else
            {
                //Debugger.Launch();
                Thread.Sleep(1000);
                Session = RootSession();
                Session = ReloadSession1(Session, "Shell_TrayWnd");
                Thread.Sleep(1000);
                var venExcel = Session.FindElementByName("Excel - 1 ventana de ejecución");
                if (venExcel != null)
                {
                    venExcel.Click();
                    Thread.Sleep(4000);
                    int count = 0;
                    while (count < 240)
                    {
                        try
                        {
                            Session = RootSession();
                            Session = ReloadSession1(Session, "XLMAIN");
                            Thread.Sleep(500);
                            var saveExcel1 = Session.FindElementsByName("Maximizar");
                            if (saveExcel1.Count > 0)
                            {
                                break;
                            }
                        }
                        catch
                        {
                            Thread.Sleep(1000);
                            count++;
                        }
                    }
                    var saveExcel = Session.FindElementsByName("Maximizar");
                    var cantidad = saveExcel.Count;
                    if (cantidad == 2)
                    {
                        saveExcel[1].Click();
                    }
                    else
                    {
                        saveExcel[0].Click();
                    }
                    //Cambio Jose
                    Session.FindElementByName("Pestaña Archivo").Click();
                    Thread.Sleep(1000);
                    //Fin Cambio Jose

                    //Session.FindElementByName("Guardar").Click();
                    Session.FindElementByName("Guardar").Click();
                    Thread.Sleep(1000);
                    Session.FindElementByName("Examinar").Click();
                    Thread.Sleep(1000);
                    Session.Keyboard.SendKeys(ruta);
                    Session.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);
                    //newFunctions_4.ScreenshotNuevo("Reporte Excel Guardado", true, file);
                    //LimpiarProcesos();
                    Thread.Sleep(2000);
                }
            }

            Thread.Sleep(2000);

        }









        //preliminar 2
        public static List<string> ReportePreliminar2(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string codProgram, string ruta )
        {
            string rutaex = @"C:\Reportes\ExportarExcel" + "_" + codProgram + "_" + Hora(); ;
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
                        newFunctions_4.ScreenshotNuevo("Preliminar 2 - Exportar excel ", true, file);
                        Thread.Sleep(400);
                        RootSession.Mouse.MouseMove(null, 730, 394);
                        RootSession.Mouse.Click(null);
                        Thread.Sleep(4500);
                        newFunctions_4.ScreenshotNuevo("Preliminar 2 - Exportar excel ", true, file);
                        Thread.Sleep(400);
                        RootSession.Mouse.MouseMove(null, 160, -3);
                        RootSession.Mouse.Click(null);
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        RootSession.Mouse.MouseMove(null, 100, -3);
                        RootSession.Mouse.Click(null);
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        newFunctions_4.ScreenshotNuevo("Preliminar 2 - Exportar excel ", true, file);
                        Thread.Sleep(400);
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(400);
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        Thread.Sleep(400);
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        Thread.Sleep(400);
                        RootSession.Keyboard.SendKeys(rutaex);
                        Thread.Sleep(1000);
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        Thread.Sleep(1400);


                        //if (bandera == "0")
                        //{
                        //    errors = newFunctions_2.generarExcel(desktopSession, file, codProgram, ruta);
                        //    if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        //}
                    

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