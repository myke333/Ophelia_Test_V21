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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_2
{
    class CrudKnmprorr : FuncionesVitales
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
                    rootSession = ReloadSession(rootSession, "TFrmTipRepo");

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
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");

            //int cont = 2;

            if (tipo == 1)
            {

                PruebaCRUD.LupaDinamica(desktopSession, crudVariables);



                //ElementList[0].Click();
                ElementList[7].SendKeys(crudVariables[1]);
                ElementList[6].SendKeys(crudVariables[2]);
                ElementList[5].SendKeys(crudVariables[3]);
                

            }
            else
            {

                ElementList[5].Clear();
                ElementList[5].SendKeys(crudVariables[4]);
            }


        }
    }
}
