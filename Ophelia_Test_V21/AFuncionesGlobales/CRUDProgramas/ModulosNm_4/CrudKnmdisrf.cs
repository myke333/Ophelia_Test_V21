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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_4
{
    class CrudKnmdisrf : FuncionesVitales
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
        
        public static List<string> ReportePreliminar(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1,int pos)
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
                List<Point> coordenadas = FuncionesGlobales.coordinatesFinder(desktopSession, 227, 117, 0);
                int coordX = coordenadas[0].X;
                int coordY = coordenadas[0].Y;
                var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
                Thread.Sleep(1000);


                try
                {
                    desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordX, coordY);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(2000);
                    rootSession = RootSession();
                    rootSession = ReloadSession(rootSession, "TFrmSemRpt");

                    var btnRad = rootSession.FindElementsByClassName("TGroupButton");
                    var elementList = rootSession.FindElementsByClassName("TEdit");
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
        public static List<string> ExpExcel(WindowsDriver<WindowsElement> desktopSession, string banExcel, string file, string codProgram, int pos, string ruta)
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
                    List<Point> coordenadas = newFunctions_2.coordinatesFinder(desktopSession);
                    int coordX = coordenadas[0].X;
                    int coordY = coordenadas[0].Y;
                    var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
                    string ClickButtonExcel = "//Pane[contains(@ClassName, 'TKNavegador')]/..";
                    var ExcelOption1 = desktopSession.FindElementsByXPath(ClickButtonExcel);

                    if (ExcelOption1.Count > 0)
                    {
                        desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordX, coordY);
                        desktopSession.Mouse.Click(null);
                        bool optionExcel = false;
                        int cont = 0;
                        rootSession = RootSession();
                        rootSession = ReloadSession(rootSession, "TFrmMenuExcel");

                        var btnRad = rootSession.FindElementsByClassName("TGroupButton");
                        var elementList = rootSession.FindElementsByClassName("TEdit");
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

                        if (banExcel == "0")
                        {
                            Thread.Sleep(1000);
                            newFunctions_2.generarExcel(rootSession, file, codProgram, ruta);
                            rootSession = RootSession();
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
            WindowsDriver<WindowsElement> rootSession = null;

            List<int> lupasOff = new List<int>() { 1 };
            List<int> lupasOff1 = new List<int>() { 0 };
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList2 = desktopSession.FindElementsByClassName("TGroupButton");



            //int cont = 2;

            if (tipo == 1)
            {
                
                PruebaCRUD.LupaDinamica(desktopSession, crudVariables);
                
                ElementList[15].SendKeys(crudVariables[1]);

            }
            else if (tipo == 2)
            {

                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "IEFrame");
                var allFrame = rootSession.FindElementsByClassName("IEFrame");
                new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2)+50 , (allFrame[0].Size.Height / 2) + 70).DoubleClick().Click().Perform();

            }
            else
            {
                ElementList[15].Clear();
                ElementList[15].SendKeys(crudVariables[2]);
            }


        }

        public static void Enter(WindowsDriver<WindowsElement> desktopSession)
        {

            WindowsDriver<WindowsElement> RootSession2 = null;
            RootSession2 = PruebaCRUD.RootSession();
            RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2000);
          
        }


        static public void reportePreliminardif(WindowsDriver<WindowsElement> desktopSession, int bandera, string file)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TGridPanel");
            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 130, 17);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);

            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession = ReloadSession(RootSession, "TFrmSemRpt");

            var btnTDVI2 = RootSession.FindElementsByClassName("TFrmSemRpt");

            if (bandera == 0)
            {
                newFunctions_4.ScreenshotNuevo("Primera opcion", true, file);
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
            else
            {
                RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 117, 127);
                RootSession.Mouse.Click(null);
                newFunctions_4.ScreenshotNuevo("SEgunda opcion", true, file);
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }


        }

        static public void reporteExceldif(WindowsDriver<WindowsElement> desktopSession, int bandera, string file)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TGridPanel");
            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 577, 17);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);

            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession = ReloadSession(RootSession, "TFrmMenuExcel");

            var btnTDVI2 = RootSession.FindElementsByClassName("TGroupButton");

            if (bandera == 0)
            {
                newFunctions_4.ScreenshotNuevo("Primera opcion", true, file);
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
            else
            {
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                newFunctions_4.ScreenshotNuevo("Segunda opcion", true, file);
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }


        }

        static public void CrudDetalle(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDetalle, string file)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");
            var btnTDVI2 = desktopSession.FindElementsByClassName("TDBGrid");

            switch (bandera)
            {
                case 0:
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 298, 99);
                    desktopSession.Mouse.Click(null);
                    newFunctions_4.ScreenshotNuevo("Ventana detalle", true, file);
                    break;
                case 1:
                    desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 97, 22);
                    desktopSession.Mouse.DoubleClick(null);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    Thread.Sleep(2000);
                    desktopSession.Keyboard.SendKeys(crudDetalle[0]);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    for (int a=0;a<4;a++)
                    {
                        desktopSession.Keyboard.SendKeys(crudDetalle[1]);
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    }
                
                    Thread.Sleep(2000);
                    desktopSession.Keyboard.SendKeys(crudDetalle[2]);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    desktopSession.Keyboard.SendKeys(crudDetalle[3]);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    desktopSession.Keyboard.SendKeys(crudDetalle[4]);
                    break;
                case 2:
                    desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 478, 29);
                    desktopSession.Mouse.DoubleClick(null);
                    desktopSession.Keyboard.SendKeys(crudDetalle[5]);
                    break;
            }


        }


    }
}
