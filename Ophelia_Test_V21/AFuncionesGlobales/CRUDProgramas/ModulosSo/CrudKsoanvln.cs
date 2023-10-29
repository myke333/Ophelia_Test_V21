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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSo
{
    class CrudKsoanvln : FuncionesVitales
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
                    rootSession = ReloadSession(rootSession, "TFrmOpcRepo");

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
                                newFunctions_4.ScreenshotNuevo("Opciones Preliminar", true, file);
                            }
                        }

                        c = c + 1;
                    }

                    var btn = rootSession.FindElementsByClassName("TButton");
                    ////Debugger.Launch();
                    foreach (var boton in btn)
                    {
                        if (boton.Text == "OK")
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

        public static void LupaDinamicaDiscriminatoria(WindowsDriver<WindowsElement> desktopSession, List<string> data, int val, List<int> lupasOffIndex = null)
        {
            List<string> errorMessages = new List<string>();
            List<Point> coordenadas = PruebaCRUD.CoordinatesFinder(desktopSession, 82, 150, 75);
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            int contLupa = 1;
            bool IsDisplayedQbe = false;
            int indexVal = val;

            if (lupasOffIndex != null)
            {
                List<Point> backupList = new List<Point>();
                for (int i = 0; i < coordenadas.Count; i++)
                {
                    if (!lupasOffIndex.Contains(i)) backupList.Add(coordenadas[i]);
                }
                coordenadas = backupList;
            }

            foreach (Point coord in coordenadas)
            {
                Thread.Sleep(2000);

                desktopSession.Mouse.MouseMove(parentElement.Coordinates, coord.X + 10, coord.Y);
                desktopSession.Mouse.Click(null);
                if (contLupa == 1)
                {
                    rootSession = RootSession();
                    rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");                   
                    var campo = rootSession.FindElementsByClassName("TEdit");
                    rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                    rootSession.Mouse.Click(null);
                    rootSession.Keyboard.SendKeys(data[indexVal]);
                    var btn = rootSession.FindElementsByClassName("TBitBtn");
                    ////Debugger.Launch();
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



        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVariables, String file)
        {
            List<string> errorMessages = new List<string>();
            int screenWidth = Screen.PrimaryScreen.Bounds.Width;
            WindowsDriver<WindowsElement> rootSession = null;

            rootSession = RootSession();
            
            List<int> lupasOff = new List<int>() { 0,1,2,3,4,5,6};
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementPage = desktopSession.FindElementsByClassName("TPageControl");


            if (tipo == 1)
            {

                {


                    ElementList[2].SendKeys(crudVariables[0]);
                    ElementList[3].SendKeys(crudVariables[1]);
                    ElementList[4].SendKeys(crudVariables[8]);
                    desktopSession.Mouse.MouseMove(ElementPage[0].Coordinates, 25, 6);
                    desktopSession.Mouse.Click(null);
                    LupaDinamicaDiscriminatoria(desktopSession, crudVariables, 2, lupasOff);
                    //newFunctions_4.findText(desktopSession, crudVariables, "TDBEdit");
                    var ElementList1 = desktopSession.FindElementsByClassName("TDBEdit");
                    ElementList1[16].SendKeys(crudVariables[5]);
                    ElementList1[15].SendKeys(crudVariables[6]);
                    ElementList1[6].Click();
                    rootSession = RootSession();
                    rootSession = ReloadSession(rootSession, "TFrmDireccion");
                    var ElementEdit = rootSession.FindElementsByClassName("TEdit");
                    var btn = rootSession.FindElementsByClassName("TBitBtn");
                    ElementEdit[4].SendKeys(crudVariables[7]);
                    btn[1].Click();
                   
                }

            }

            else
            {

                
                ElementList[17].Clear();
                ElementList[17].SendKeys(crudVariables[9]);

            }
        }
    }
}
