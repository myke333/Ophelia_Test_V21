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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosPe
{
    class CrudKpenotic : FuncionesVitales
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
        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TGroupButton");

            if(tipo==1)
            {
                foreach (var elem in ElementList)
                {
                    if (elem.Text == "Acceso Smart People")
                    {
                        elem.Click();
                    }
                }
                lapizEdit(desktopSession);
                Thread.Sleep(2000);
                WindowsDriver<WindowsElement> rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TFrmRichEditor");
                var Campo = rootSession.FindElementsByClassName("TDBRichEdit");
                Campo[0].SendKeys(data[0]);
                var cerrar = rootSession.FindElementByName("Cerrar");
                cerrar.Click();
                newFunctions_4.ScreenshotNuevo("Agregar Registro", true, file);
            }
            else
            {
                foreach (var elem in ElementList)
                {
                    if (elem.Text == "Reclutamiento Web")
                    {
                        elem.Click();
                    }
                }
                newFunctions_4.ScreenshotNuevo("Editar Registro", true, file);
            }
        }
        private static void lapizEdit(WindowsDriver<WindowsElement> desktopSession)
        {
            List<Point> coordenadas = coordinatesFinder2(desktopSession);
            int coordx = coordenadas[0].X;
            int coordy = coordenadas[0].Y;
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordx, coordy);
            desktopSession.Mouse.Click(null);
        }

        private static List<Point> coordinatesFinder2(WindowsDriver<WindowsElement> session)
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

            for (int x = 0; x < bmpSource.Width; x++)
            {
                for (int y = 0; y < bmpSource.Height; y++)
                {
                    Color clrPixel = bmpSource.GetPixel(x, y);
                    
                    if (clrPixel.R == 216 && clrPixel.G == 174 && clrPixel.B == 97)
                    {
                        puntos.X = x;
                        puntos.Y = y;
                        if (x > lastx + 20 || y > lasty + 20)
                        {
                            coordenadas.Add(puntos);
                            lastx = x;
                            lasty = y;
                        }
                    }
                }
            }
            return coordenadas;
        }
    }
}
