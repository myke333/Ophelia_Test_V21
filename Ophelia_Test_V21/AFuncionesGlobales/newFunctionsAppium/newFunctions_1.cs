
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
using System.Drawing.Imaging;
using System.IO;


namespace Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium
{
    class newFunctions_1:FuncionesVitales
    {
        static public WindowsDriver<WindowsElement> RootSession()
        {

            var appCapabilities = new AppiumOptions();
            appCapabilities.AddAdditionalCapability("app", "Root");
            var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
            return ApplicationSession;
        }
        static public WindowsDriver<WindowsElement> ReloadSession(WindowsDriver<WindowsElement> session,string parametro)
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
       
        public static void lapiz(WindowsDriver<WindowsElement> desktopSession)
        {
            List<Point> coordenadas = coordinatesFinder2(desktopSession);
            int coordx = coordenadas[0].X;
            int coordy = coordenadas[0].Y;
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordx, coordy);
            desktopSession.Mouse.Click(null);
        }
        /// <summary>
        /// La bandera 0, cuando no tiene sumatoria
        /// La bandera 1 cuando tiene Sumatoria
        /// </summary>
        /// <param name="desktopSession"></param>
        public static List<string> Sumatory(WindowsDriver<WindowsElement> desktopSession, string BanderaSum, string file,string mover=null)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            ////Debugger.Launch();
            int screenWidth = Screen.PrimaryScreen.Bounds.Width;
            int Height = Screen.PrimaryScreen.Bounds.Height;

            if (BanderaSum == "1")
            {
                //var allWindowHandles = rootSession.WindowHandles;
                //WebDriverWait wait = new WebDriverWait(rootSession, TimeSpan.FromSeconds(1));

                bool validar = coordinatesFinder(desktopSession, 240, 240, 240, 1550, 330, 2);
                if (validar == true)
                {
                    lapiz(desktopSession);
                }

                int counterTry = 0;
                ////Debugger.Launch();
                //var fieldInt = rootSession.FindElementsByClassName("TKactusDBgrid");
                //new Actions(rootSession).MoveToElement(fieldInt[0], 0, 0).MoveByOffset(50, 40).ClickAndHold().Click().Perform();
                var TDBGridInplaceEdit = desktopSession.FindElementsByClassName("TDBGridInplaceEdit");
                TDBGridInplaceEdit[0].Click();
                Thread.Sleep(2000);
                new Actions(rootSession).Click().Perform();
                int contscreen = 0;
                if (mover != null)
                {
                    int m = Convert.ToInt32(mover);
                    for (int i = 1; i <= m; i++)
                    {
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                    }
                }
                while (counterTry != 10 && counterTry != 20)
                {
                    try
                    {
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Alt + "H" + OpenQA.Selenium.Keys.Alt);

                        for (int i = 1; i <= 4; i++)
                        {
                            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                            Thread.Sleep(100);
                        }

                        //PRIMER SCREEN DEL MENU SUMATORIA
                        if (contscreen == 0) { newFunctions_4.ScreenshotNuevo("Menú Sumatoria", true, file); ; contscreen = 1; }
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        Thread.Sleep(2000);

                        bool ventana = coordinatesFinder(desktopSession, 27, 101, 206, screenWidth / 2, 347, 1);
                        rootSession = RootSession();
                        string navigator = AbrirPrograma.SelectNavigator();
                        if (navigator == "IE")
                        {
                            rootSession = ReloadSession(rootSession, "IEFrame"); 
                            var allFrame = rootSession.FindElementsByClassName("IEFrame");
                            Thread.Sleep(1000);
                            new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset(allFrame[0].Size.Width / 2, allFrame[0].Size.Height / 2).DoubleClick().Click().Perform();
                            rootSession = RootSession();
                        }
                        else
                        {
                            //rootSession = ReloadSession(rootSession, "Chrome_WidgetWin_1");
                            var allFrame = rootSession.FindElementsByClassName("TKactusDBgrid");
                            Thread.Sleep(1000);
                            //Debugger.Launch();
                            new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset(allFrame[0].Size.Width / 2, allFrame[0].Size.Height / 2).DoubleClick().Click().Perform();
                            rootSession = RootSession();
                        }

                        if (ventana == false)
                        {
                            if (navigator == "IE")
                            {
                                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                Thread.Sleep(2000);
                                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                                counterTry += 1;
                            }
                            else
                            {
                                var valsum = desktopSession.FindElementsByClassName("TKactusDBgrid");
                                //var valsum = rootSession.FindElementByXPath("//*[contains(@Name, 'Actividades')]");
                                desktopSession.Mouse.MouseMove(valsum[0].Coordinates, 771, 220);
                                Thread.Sleep(500);
                                desktopSession.Mouse.Click(null);
                                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                Thread.Sleep(2000);
                                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                                counterTry += 1;
                            }

                        }
                        else
                        {
                            if (navigator == "IE")
                            {
                                newFunctions_4.ScreenshotNuevo("Valor Sumatoria", true, file);
                                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                counterTry = 20;
                                break;
                            }
                            else
                            {
                                newFunctions_4.ScreenshotNuevo("Valor Sumatoria", true, file);
                                Thread.Sleep(1000);
                                rootSession = RootSession(); //TKactusDBgrid
                                var valsum = desktopSession.FindElementsByClassName("TKactusDBgrid");
                                //var valsum = rootSession.FindElementByXPath("//*[contains(@Name, 'Actividades')]");
                                desktopSession.Mouse.MouseMove(valsum[0].Coordinates, 771, 220);
                                Thread.Sleep(500);
                                desktopSession.Mouse.Click(null);
                                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                counterTry = 20;
                                break;
                            }

                        }
                    }
                    catch (Exception)
                    {
                        errorMessages.Add("Se ha encontrado una Excepción y se sale el try");
                        break;
                    }
                    if (counterTry==10)
                    {
                        errorMessages.Add("No se han encontrado campos enteros para la Sumatoria");
                    }
                }
            }
            return errorMessages;
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
                    
                    if (clrPixel.R == 233 && clrPixel.G == 197 && clrPixel.B == 0)
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

        public static bool coordinatesFinder(WindowsDriver<WindowsElement> session,int R, int G, int B, int fil, int col,int tipo)
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
            else if(tipo==1)
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

        public static void ScreenshotNuevoSP105(string maestro,string database)
        {
            try
            {
                //Creating a new Bitmap object

                int screenWidth = Screen.PrimaryScreen.Bounds.Width;
                int screenHeight = Screen.PrimaryScreen.Bounds.Height;

                string name = maestro;
                string Path = string.Format(@"C:\SP105\{0}\", database);
                if (!Directory.Exists(Path))
                {
                    Directory.CreateDirectory(Path);
                }

                Bitmap captureBitmap = new Bitmap(screenWidth, screenHeight, PixelFormat.Format32bppArgb);

                Rectangle captureRectangle = Screen.AllScreens[0].Bounds;

                Graphics captureGraphics = Graphics.FromImage(captureBitmap);

                captureGraphics.CopyFromScreen(captureRectangle.Left, captureRectangle.Top, 0, 0, captureRectangle.Size);

                captureBitmap = (new Bitmap(captureBitmap, new Size(1600, 900)));

                captureBitmap.Save(string.Format(Path +"{0}.bmp", name), ImageFormat.Bmp);

            }

            catch (Exception ex)

            {

                MessageBox.Show(ex.Message);

            }
        }
    }
}
