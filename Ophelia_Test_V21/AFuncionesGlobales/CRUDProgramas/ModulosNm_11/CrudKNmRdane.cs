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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_11
{
	class CrudKNmRdane:FuncionesVitales
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

        


        static public void CrudRDane(WindowsDriver<WindowsElement> desktopSession, string tipo)
		{
            var Elementlist = desktopSession.FindElementsByClassName("TEdit");

            switch (tipo)
            {
                case ("0"):
                    desktopSession.Mouse.MouseMove(Elementlist[3].Coordinates, 75, 5);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> FecPro = null;
                    FecPro = PruebaCRUD.RootSession();
                    FecPro.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[3].Coordinates, 190, 5);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    WindowsDriver<WindowsElement> FecIni = null;
                    FecIni = PruebaCRUD.RootSession();
                    FecIni.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[3].Coordinates, 300, 5);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    WindowsDriver<WindowsElement> FecFin = null;
                    FecFin = PruebaCRUD.RootSession();
                    FecFin.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[3].Coordinates, 75, -36);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[3].Coordinates, 190, -36);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[3].Coordinates, 300, -36);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[3].Coordinates, 190, -100);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[3].Coordinates, 20, -60);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> Datos1 = null;
                    Datos1 = PruebaCRUD.RootSession();
                    Datos1.Keyboard.SendKeys("1");

                    desktopSession.Mouse.MouseMove(Elementlist[3].Coordinates, 20, -20);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> Datos2 = null;
                    Datos2 = PruebaCRUD.RootSession();
                    Datos2.Keyboard.SendKeys("100");
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[3].Coordinates, 400, -60);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> Datos3 = null;
                    Datos3 = PruebaCRUD.RootSession();
                    Datos3.Keyboard.SendKeys("110");
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[3].Coordinates, 400, -20);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> Datos4 = null;
                    Datos4 = PruebaCRUD.RootSession();
                    Datos4.Keyboard.SendKeys("102");
                    Thread.Sleep(1000);

                    desktopSession.Mouse.MouseMove(Elementlist[3].Coordinates, 400, 20);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    WindowsDriver<WindowsElement> Datos5 = null;
                    Datos5 = PruebaCRUD.RootSession();
                    Datos5.Keyboard.SendKeys("200");
                    Thread.Sleep(1000);

                    break;
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
                        newFunctions_4.ScreenshotNuevo("Preliminar ", true, file);
                        

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




    }
}
