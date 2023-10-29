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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosAc
{
    class CrudKacmaans : FuncionesVitales
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




        static public void AgregarCrud(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal)
        {

            //var btnTDVI = desktopSession.FindElementsByClassName("TGroupBox");
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");
            //var btnTDVI = desktopSession.FindElementsByClassName("TPanel");
            //Debugger.Launch();

            var btnTDVI = desktopSession.FindElementsByClassName("TDBEdit");

            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 628, 194);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[2]);
                Thread.Sleep(2000);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 730, 194);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
                Thread.Sleep(2000);

                btnTDVI[5].SendKeys(crudPrincipal[1]);

            }
            else
            {
                btnTDVI[5].Clear();
                btnTDVI[5].SendKeys(crudPrincipal[4]);
            }
        }



        static public void Click(WindowsDriver<WindowsElement> desktopSession)
        {
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");

            desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 636, 33);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1500);

            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession = ReloadSession(RootSession, "TKQBE");

            var btnTDVI = RootSession.FindElementsByClassName("TTabSheet");

            RootSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 600, 16);
            RootSession.Mouse.Click(null);
            Thread.Sleep(2000);

            WindowsDriver<WindowsElement> RootSession3 = null;
            RootSession3 = PruebaCRUD.RootSession();
            RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(3000);

            WindowsDriver<WindowsElement> RootSession2 = null;
            RootSession2 = PruebaCRUD.RootSession();
            RootSession2 = ReloadSession(RootSession2, "TCaptura");

            var btnTDVI3 = RootSession2.FindElementsByClassName("TKactusDBgrid");

            RootSession2.Mouse.MouseMove(btnTDVI3[0].Coordinates, 235, 153);
            RootSession2.Mouse.Click(null);
            RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2000);
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
                        Thread.Sleep(800);

                        WindowsDriver<WindowsElement> RootSession = null;
                        RootSession = PruebaCRUD.RootSession();
                        RootSession = ReloadSession(RootSession, "TFrmReporte");
                        newFunctions_4.ScreenshotNuevo("Preliminar 1 - Estandar - Anual -  Reporte", true, file);
                        Thread.Sleep(500);


                        // Click en aceptar.
                        var btnTDVI2 = RootSession.FindElementsByClassName("TFrmReporte");

                        RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 123, 245);
                        RootSession.Mouse.Click(null);
                        Thread.Sleep(9000);                        
                       


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
        static public void ReportePreliminar2(WindowsDriver<WindowsElement> desktopSession, string file, string pdf1)
        {

            var btnTDVI23 = desktopSession.FindElementsByClassName("TPanel");

            desktopSession.Mouse.MouseMove(btnTDVI23[0].Coordinates, 120, -50);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1500);

            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession = ReloadSession(RootSession, "TFrmReporte");
            Thread.Sleep(1400);


            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            Thread.Sleep(400);
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
            Thread.Sleep(400);

            newFunctions_4.ScreenshotNuevo("Preliminar 2 - Estandar - Anual - Archivo Excel", true, file);
            Thread.Sleep(1500);


            WindowsDriver<WindowsElement> RootSession2 = null;
            RootSession2 = PruebaCRUD.RootSession();

            RootSession2.Keyboard.SendKeys(pdf1);
            Thread.Sleep(1000);

            RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(3500);



            // Click en aceptar.
            var btnTDVI2 = RootSession.FindElementsByClassName("TFrmReporte");

            RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 123, 245);
            RootSession.Mouse.Click(null);
            Thread.Sleep(15000);


            newFunctions_4.ScreenshotNuevo("Preliminar 2 - Estandar - Anual - Estado Archivo Excel", true, file);
            Thread.Sleep(500);


            RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(500);

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
                        Thread.Sleep(800);

                        WindowsDriver<WindowsElement> RootSession = null;
                        RootSession = PruebaCRUD.RootSession();
                        RootSession = ReloadSession(RootSession, "TFrmReporte");                                                
                        var btnTDVI2 = RootSession.FindElementsByClassName("TFrmReporte");


                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                        Thread.Sleep(400);

                        newFunctions_4.ScreenshotNuevo("Preliminar 3 - Estandar - Mensual -  Reporte", true, file);
                        Thread.Sleep(500);

                        //Clik en Aceptar

                        RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 123, 245);
                        RootSession.Mouse.Click(null);
                        Thread.Sleep(9000);



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



        //preliminar 4
        static public void ReportePreliminar4(WindowsDriver<WindowsElement> desktopSession, string file, string pdf1)
        {

            var btnTDVI23 = desktopSession.FindElementsByClassName("TPanel");

            desktopSession.Mouse.MouseMove(btnTDVI23[0].Coordinates, 120, -50);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1500);


            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession = ReloadSession(RootSession, "TFrmReporte");
            Thread.Sleep(1400);


            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
            Thread.Sleep(400);
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            Thread.Sleep(400);
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
            Thread.Sleep(400);

            newFunctions_4.ScreenshotNuevo("Preliminar 4 - Estandar - Mensual - Archivo Excel", true, file);
            Thread.Sleep(1500);


            //Se abre ventana de exportar archivo 
            WindowsDriver<WindowsElement> RootSession2 = null;
            RootSession2 = PruebaCRUD.RootSession();


            //se pasa nombre del archivo para guardarlo
            RootSession2.Keyboard.SendKeys(pdf1);
            Thread.Sleep(1000);

            RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(3500);


            // Click en aceptar.
            var btnTDVI2 = RootSession.FindElementsByClassName("TFrmReporte");

            RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 123, 245);
            RootSession.Mouse.Click(null);
            Thread.Sleep(15000);

            newFunctions_4.ScreenshotNuevo("Preliminar 4 - Estandar - Anual - Estado del proceso de guardar Archivo Excel", true, file);
            Thread.Sleep(500);

            WindowsDriver<WindowsElement> RootSession3 = null;
            RootSession2 = PruebaCRUD.RootSession();

            RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2500);
         
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
                        Thread.Sleep(800);


                        WindowsDriver<WindowsElement> RootSession = null;
                        RootSession = PruebaCRUD.RootSession();
                        RootSession = ReloadSession(RootSession, "TFrmReporte");
                        

                        // Se escoge preliminar.
                        var btnTDVI2 = RootSession.FindElementsByClassName("TFrmReporte");

                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(800);
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(800);
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                        Thread.Sleep(800);


                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(400);
                        RootSession.Keyboard.SendKeys("1");
                        Thread.Sleep(1000);
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(400);
                        RootSession.Keyboard.SendKeys("10");
                        Thread.Sleep(1000);
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(400);
                        RootSession.Keyboard.SendKeys("2022");
                        Thread.Sleep(1000);
                        RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(400);
                        RootSession.Keyboard.SendKeys("31/10/2022");
                        Thread.Sleep(1000);

                        // Screamshot de la configuracion de datos para el preliminar 5
                        newFunctions_4.ScreenshotNuevo("Preliminar 5 - Ajuste Salarial ", true, file);
                        Thread.Sleep(500);

                        // Click en aceptar.
                        RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 123, 245);
                        RootSession.Mouse.Click(null);
                        Thread.Sleep(10000);



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


        //preliminar 6
        static public void ReportePreliminar6(WindowsDriver<WindowsElement> desktopSession, string file, string pdf1)
        {

            var btnTDVI23 = desktopSession.FindElementsByClassName("TPanel");

            desktopSession.Mouse.MouseMove(btnTDVI23[0].Coordinates, 120, -50);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1500);


            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession = ReloadSession(RootSession, "TFrmReporte");
            Thread.Sleep(1400);


            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            Thread.Sleep(400);
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            Thread.Sleep(400);
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
            Thread.Sleep(400);
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
            Thread.Sleep(400);


            newFunctions_4.ScreenshotNuevo("Preliminar 6 - Graficas - Dispercion de datos", true, file);
            Thread.Sleep(1500);


            // Click en aceptar.
            var btnTDVI2 = RootSession.FindElementsByClassName("TFrmReporte");

            RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 123, 245);
            RootSession.Mouse.Click(null);
            Thread.Sleep(15000);

            newFunctions_4.ScreenshotNuevo("Preliminar 6 Ventana - Graficas - Dispercion de datos", true, file);
            Thread.Sleep(500);


            WindowsDriver<WindowsElement> RootSession2 = null;
            RootSession2 = PruebaCRUD.RootSession();

            RootSession2.Keyboard.SendKeys(pdf1);
            Thread.Sleep(1000);

            RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(3500);



            WindowsDriver<WindowsElement> RootSession3 = null;
            RootSession2 = PruebaCRUD.RootSession();

            RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(500);

        }



        //preliminar 6.2 / 7
        static public void ReportePreliminar7(WindowsDriver<WindowsElement> desktopSession, string file, string pdf1)
        {

            var btnTDVI23 = desktopSession.FindElementsByClassName("TPanel");

            desktopSession.Mouse.MouseMove(btnTDVI23[0].Coordinates, 120, -50);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1500);

            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession = ReloadSession(RootSession, "TFrmReporte");
            Thread.Sleep(1400);
            var btnTDVI2 = RootSession.FindElementsByClassName("TFrmReporte");


            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            Thread.Sleep(400);
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            Thread.Sleep(400);
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
            Thread.Sleep(400);
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
            Thread.Sleep(400);

            //Click para grafica2
            RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 40, 140);
            RootSession.Mouse.Click(null);
            Thread.Sleep(500);

            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
            Thread.Sleep(400);


            newFunctions_4.ScreenshotNuevo("Preliminar 6.2 - Graficas - Tendencia Central", true, file);
            Thread.Sleep(1500);


            // Click en aceptar.            
            RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 123, 245);
            RootSession.Mouse.Click(null);
            Thread.Sleep(15000);

            newFunctions_4.ScreenshotNuevo("Preliminar 7 Ventana - Graficas - Tendencia Central", true, file);
            Thread.Sleep(500);


            WindowsDriver<WindowsElement> RootSession2 = null;
            RootSession2 = PruebaCRUD.RootSession();

            RootSession2.Keyboard.SendKeys(pdf1);
            Thread.Sleep(1000);

            RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(3500);



            WindowsDriver<WindowsElement> RootSession3 = null;
            RootSession2 = PruebaCRUD.RootSession();

            RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(500);

        }


        //preliminar 6.3 / 8
        static public void ReportePreliminar8(WindowsDriver<WindowsElement> desktopSession, string file, string pdf1)
        {

            var btnTDVI23 = desktopSession.FindElementsByClassName("TPanel");

            desktopSession.Mouse.MouseMove(btnTDVI23[0].Coordinates, 120, -50);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1500);

            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession = ReloadSession(RootSession, "TFrmReporte");
            Thread.Sleep(1400);
            var btnTDVI2 = RootSession.FindElementsByClassName("TFrmReporte");


            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            Thread.Sleep(400);
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            Thread.Sleep(400);
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
            Thread.Sleep(400);
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
            Thread.Sleep(400);

            //Click para grafica2
            RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 40, 140);
            RootSession.Mouse.Click(null);
            Thread.Sleep(500);

            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
            Thread.Sleep(400);

            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
            Thread.Sleep(400);


            newFunctions_4.ScreenshotNuevo("Preliminar 6.3 - Graficas - Banda de equidad interna", true, file);
            Thread.Sleep(1500);


            // Click en aceptar.            
            RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 123, 245);
            RootSession.Mouse.Click(null);
            Thread.Sleep(15000);

            newFunctions_4.ScreenshotNuevo("Preliminar 8 Ventana - Graficas - Banda de equidad interna", true, file);
            Thread.Sleep(500);


            WindowsDriver<WindowsElement> RootSession2 = null;
            RootSession2 = PruebaCRUD.RootSession();

            RootSession2.Keyboard.SendKeys(pdf1);
            Thread.Sleep(1000);

            RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(3500);



            WindowsDriver<WindowsElement> RootSession3 = null;
            RootSession2 = PruebaCRUD.RootSession();

            RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(500);

        }

        //preliminar 6.4 / 9
        static public void ReportePreliminar9(WindowsDriver<WindowsElement> desktopSession, string file, string pdf1)
        {

            var btnTDVI23 = desktopSession.FindElementsByClassName("TPanel");

            desktopSession.Mouse.MouseMove(btnTDVI23[0].Coordinates, 120, -50);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1500);

            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession = ReloadSession(RootSession, "TFrmReporte");
            Thread.Sleep(1400);
            var btnTDVI2 = RootSession.FindElementsByClassName("TFrmReporte");


            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            Thread.Sleep(400);
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            Thread.Sleep(400);
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
            Thread.Sleep(400);
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
            Thread.Sleep(400);

            //Click para grafica2
            RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 40, 140);
            RootSession.Mouse.Click(null);
            Thread.Sleep(500);

            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
            Thread.Sleep(400);
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
            Thread.Sleep(400);
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
            Thread.Sleep(400);



            newFunctions_4.ScreenshotNuevo("Preliminar 6.4 - Graficas - Valores de mercado ", true, file);
            Thread.Sleep(1500);


            // Click en aceptar.            
            RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 123, 245);
            RootSession.Mouse.Click(null);
            Thread.Sleep(15000);

            newFunctions_4.ScreenshotNuevo("Preliminar 9 Ventana - Graficas - Valores de mercado ", true, file);
            Thread.Sleep(500);


            WindowsDriver<WindowsElement> RootSession2 = null;
            RootSession2 = PruebaCRUD.RootSession();

            RootSession2.Keyboard.SendKeys(pdf1);
            Thread.Sleep(1000);

            RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(3500);



            WindowsDriver<WindowsElement> RootSession3 = null;
            RootSession2 = PruebaCRUD.RootSession();

            RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(500);

        }


    }
}

