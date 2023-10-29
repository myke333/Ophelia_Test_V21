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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_4
{
    class CrudKnmmovse : FuncionesVitales
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

                return ApplicationSession;
            }
            catch (Exception ex) { return null; }
        }
        public static List<string> ReportePreliminar(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1, int pos, string nomPrograma = null)
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
                List<Point> coordenadas =  FuncionesGlobales.coordinatesFinder(desktopSession, 227, 117, 0);
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
                    rootSession = ReloadSession(rootSession, "TFrmUImpRep");

                    var btnRad = rootSession.FindElementsByClassName("TGroupButton");
                   
                    int c = 0;
                    ////Debugger.Launch();
                    foreach (var btRd in btnRad)
                    {
                        if (pos == c)
                        {
                            
                            {
                                btRd.Click();
                                newFunctions_4.ScreenshotNuevo("Opcion Prelimiar", true, file);
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
            return errorMessages;
        }
        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVariables, String file)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            List<int> lupasOff = new List<int>() { 3,4,5,6 };
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            
            if (tipo == 1)
            {
                newFunctions_4.LupaDinamicaDiscriminatoria(desktopSession,crudVariables,lupasOff);
                ElementList[27].SendKeys(crudVariables[3]);
                ElementList[15].SendKeys(crudVariables[4]);

            }
            else
            {
                ElementList[15].Clear();
                ElementList[15].SendKeys(crudVariables[5]);
            }




        }
    }
}
