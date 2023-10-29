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
using System.IO;
using System.Windows.Forms;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSo
{
    class CrudKsohicli : FuncionesVitales
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
        public static List<Point> coordinatesFinder(WindowsDriver<WindowsElement> session)
        {
            Screenshot image = ((ITakesScreenshot)session).GetScreenshot();
            string name = "ProgramaExcel";
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
                    if (clrPixel.R == 151 && clrPixel.G == 179 && clrPixel.B == 143)
                    {
                        puntos.X = x;
                        puntos.Y = y;
                        if (x > lastx + 10 || y > lasty + 10)
                        {
                            coordenadas.Add(puntos);
                            lastx = x;
                            lasty = y;
                            romper = 1;
                            break;
                        }
                    }
                }
                if (romper == 1)
                {
                    break;
                }
            }
            return coordenadas;
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

        public static List<string> QBEQry(WindowsDriver<WindowsElement> desktopSession, string bandera, string QbeFilter, string file)
        {
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
                    List<Point> coordenadas = coordinatesFinder(desktopSession, 244, 202, 32);

                    int coordX = coordenadas[0].X;
                    int coordY = coordenadas[0].Y;
                    var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));

                    bool IsDisplayedQbe = false;


                    desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordX, coordY);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    newFunctions_4.ScreenshotNuevo("QBE", true, file);
                    rootSession = RootSession();
                    rootSession = ReloadSession(rootSession, "TKQBE");
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
                    if (QbeFilter!="0")
                        {
                            elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/Pane[contains(@ClassName, 'TStringGrid')]/*"));
                            elements[0].SendKeys(QbeFilter);
                        }
                        elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                        rootSession.Mouse.MouseMove(elements[8].Coordinates, 20, 20);
                        rootSession.Mouse.Click(null);
                        Thread.Sleep(500);

                        if (QbeFilter=="0")
                        {
                            rootSession = ReloadSession(rootSession, "TKQBE");
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
        public static List<string> ExpExcel(WindowsDriver<WindowsElement> desktopSession, string banExcel, string file, string codProgram, string ruta, int pos)
        {
            List<string> errorMessages = new List<string>();
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            try
            {
                wait.Until(driver =>
                {
                    WindowsDriver<WindowsElement> rootSession = null;
                    List<Point> coordenadas = coordinatesFinder(desktopSession);
                    int coordX = coordenadas[0].X;
                    int coordY = coordenadas[0].Y;
                    var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
                    string ClickButtonExcel = "//Pane[contains(@ClassName, 'TKNavegador')]/..";
                    var ExcelOption1 = desktopSession.FindElementsByXPath(ClickButtonExcel);

                    

                    rootSession = RootSession();
                    rootSession = ReloadSession(rootSession, "TFrmMenExpo");

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
                   
                    foreach (var boton in btn)
                    {
                        if (boton.Text == "Aceptar")
                        {
                            boton.Click();

                        }
                    }
                    if (ExcelOption1.Count > 0)
                    {
                        desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordX, coordY);
                        desktopSession.Mouse.Click(null);
                        bool optionExcel = false;
                        int cont = 0;
                        if (banExcel == "0")
                        {
                            Thread.Sleep(1000);
                            newFunctions_2.generarExcel(rootSession, file, codProgram, ruta);
                            rootSession = RootSession();
                        }
                        else if (banExcel == "1")
                        {
                            rootSession = RootSession();
                            rootSession = ReloadSession(rootSession, "TForm");
                        }
                        else if (banExcel == "2")
                        {
                            rootSession = RootSession();
                            rootSession = ReloadSession(rootSession, "TFrmMenuExcel");
                        }
                        else if (banExcel == "3")
                        {

                            rootSession = RootSession();
                            rootSession = ReloadSession(rootSession, "TFrmMenExpo");
                        }

                        if (banExcel == "1" || banExcel == "2" || banExcel == "3")
                        {
                            newFunctions_4.ScreenshotNuevo("Opciones Exportar Excel", true, file);
                            var contButton = rootSession.FindElementsByClassName("TGroupButton");
                            int cont2 = contButton.Count() - 1;
                            rootSession.FindElementsByClassName("TGroupButton");
                            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                            Thread.Sleep(500);
                            newFunctions_2.generarExcel(rootSession, file, codProgram, ruta);
                            int i;
                            for (i = cont2 - 1; i > -1; i--)
                            {
                                desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordX, coordY);
                                desktopSession.Mouse.Click(null);
                                rootSession.FindElementsByClassName("TGroupButton")[cont2 - 1].Click();
                                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                newFunctions_2.generarExcel(rootSession, file, codProgram, ruta);
                                cont2 = cont2 - 1;
                            }
                        }

                    }
                    return rootSession != null;
                });
            }

            catch
            {
                Debug.WriteLine("No se encuentra la opcion de excel para este programa o no hay registros para exportar");
            }
            string pathImg = "C:\\imagenesTest\\Programa.Png";
            if (File.Exists(pathImg))
            {
                File.Delete(pathImg);
            }
            return errorMessages;
        }




        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVariables, String file)
        {
            List<string> errorMessages = new List<string>();
            int screenWidth = Screen.PrimaryScreen.Bounds.Width;
            WindowsDriver<WindowsElement> rootSession = null;
            List<int> lupasOff = new List<int>() { 1,2,3,4 };
            rootSession = RootSession();
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementButton = desktopSession.FindElementsByClassName("TDBComboBox");


            if (tipo == 1)
            {

                {

                    newFunctions_4.LupaDinamicaDiscriminatoria(desktopSession, crudVariables, lupasOff);
                    ElementList[2].SendKeys(crudVariables[1]);
                    ElementButton[2].SendKeys(crudVariables[2]);

                     
                }

            }

            else
            {

                ElementButton[2].Clear();
                ElementButton[2].SendKeys(crudVariables[3]);

            }
        }
       
    }
}
