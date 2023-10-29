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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBi
{
    class CrudKbiprodi : FuncionesVitales
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

    public static void LupaDinamica(WindowsDriver<WindowsElement> desktopSession, List<string> data)
    {
        List<string> errorMessages = new List<string>();
        List<Point> coordenadas = CoordinatesFinder(desktopSession, 82, 150, 75);
        WindowsDriver<WindowsElement> rootSession = null;
        rootSession = RootSession();
        var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
        int contLupa = 1;
        bool IsDisplayedQbe = false;
        int indexVal = 0;
        foreach (Point coord in coordenadas)
        {
            Thread.Sleep(2000);
            //Actions noteClicks = new Actions(desktopSession);
            //noteClicks.MoveToElement(parentElement).MoveByOffset(coord.X + 10, coord.Y).Click().Perform();

            desktopSession.Mouse.MouseMove(parentElement.Coordinates, coord.X + 10, coord.Y);
            desktopSession.Mouse.Click(null);
            if (contLupa == 1)
            {
                try
                {
                    rootSession = RootSession();
                    rootSession = ReloadSession(rootSession, "TKQBE");
                }
                catch
                {

                }
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
                    if (!string.IsNullOrEmpty(data[indexVal]))
                    {
                        elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/Pane[contains(@ClassName, 'TStringGrid')]/*"));
                        elements[0].SendKeys(data[indexVal]);
                    }
                    elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                    new Actions(rootSession).MoveToElement(elements[8], 0, 0).MoveByOffset(20, 10).DoubleClick().Perform();
                    Thread.Sleep(500);
                    IsDisplayedQbe = false;

                }

                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TCaptura");
                //var campo = rootSession.FindElementsByXPath("//Window[@ClassName=\"TCaptura\"][@Name=\" LISTA DE EMPLEADOS\"]/Edit[@ClassName=\"TEdit\"]");
                var campo = rootSession.FindElementsByClassName("TEdit");
                ////Debugger.Launch();
                rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                rootSession.Mouse.Click(null);
                rootSession.Keyboard.SendKeys(data[indexVal]);
                //Aqui se puede enviar un valor

                var btn = rootSession.FindElementsByClassName("TBitBtn");

                foreach (var boton in btn)
                {
                    if (boton.Text == "Aceptar")
                    {
                        rootSession.Mouse.MouseMove(boton.Coordinates, 10, 10);
                        rootSession.Mouse.Click(null);
                    }
                }
                indexVal++;

            }
        }

    }
        public static List<string> ReportePreliminar(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1, int pos)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();

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
                    rootSession = ReloadSession(rootSession, "TFrmImpRep");

                    var btnRad = rootSession.FindElementsByClassName("TGroupButton");
                    int c = 0;
                    ////Debugger.Launch();
                    foreach (var btRd in btnRad)
                    {
                        if (pos == c)
                        {
                            btRd.Click();

                        }
                        c = c + 1;
                    }

                    var btn = rootSession.FindElementsByClassName("TBitBtn");
                    ////Debugger.Launch();
                    foreach (var boton in btn)
                    {
                        if (boton.Text == "OK")
                        {
                            boton.Click();
                        }
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
        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVariables, String file)
    {
        List<string> errorMessages = new List<string>();

        var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
        var ElementComboBox = desktopSession.FindElementsByClassName("TDBComboBox");

       if (tipo == 1)
        {

            {
                LupaDinamica(desktopSession, crudVariables);
                ElementList[24].SendKeys(crudVariables[1]);
                ElementComboBox[0].SendKeys(crudVariables[2]);

            }

        }
        else
        {
            ElementComboBox[0].SendKeys(crudVariables[3]);
        }
    }
    public static List<Point> CoordinatesFinder(WindowsDriver<WindowsElement> session, int R, int G, int B)
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
    public static void cerrarBorrar(WindowsDriver<WindowsElement> desktopSession)
    {
        WindowsDriver<WindowsElement> rootSession = null;
        rootSession = RootSession();
        rootSession = ReloadSession(rootSession, "IEFrame");
        var allFrame = rootSession.FindElementsByClassName("IEFrame");
        new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset(allFrame[0].Size.Width / 2, allFrame[0].Size.Height / 2).DoubleClick().Click().Perform();
        rootSession = RootSession();
        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
    }
}
}
