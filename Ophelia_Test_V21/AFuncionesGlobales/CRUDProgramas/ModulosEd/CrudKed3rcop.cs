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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosEd
{
    class CrudKed3rcop : FuncionesVitales
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

        public static List<string> QBEQry(WindowsDriver<WindowsElement> desktopSession, string bandera, string QbeFilter, string file)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            bool IsDisplayedQbe = false;
            if (bandera == "0")
            {


                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "IEFrame");
                var allFrame = rootSession.FindElementsByClassName("IEFrame");
                new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) + 80, (allFrame[0].Size.Height / 2) + 110).DoubleClick().Click().Perform();
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
                    if (QbeFilter != "0")
                    {
                        elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/Pane[contains(@ClassName, 'TStringGrid')]/*"));
                        elements[0].SendKeys(QbeFilter);
                    }
                    elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                    rootSession.Mouse.MouseMove(elements[8].Coordinates, 20, 20);
                    rootSession.Mouse.Click(null);
                    Thread.Sleep(500);

                    if (QbeFilter == "0")
                    {
                        rootSession = ReloadSession(rootSession, "TKQBE");
                        elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("//Button[contains(@Name, 'Sí')]"));
                        rootSession.Mouse.MouseMove(elements[0].Coordinates, 20, 20);
                        rootSession.Mouse.Click(null);
                    }
                    Thread.Sleep(500);
                    newFunctions_4.ScreenshotNuevo("Registros QBE", true, file);
                }


            }
            return errorMessages;
        }

        public static List<string> ReportePreliminar(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1,int pos, string QbeFilter)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            WindowsElement codPanel = null;
            bool brk = false;
            bool IsDisplayedPreWin = false;

            try
            {
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "IEFrame");
                var allFrame = rootSession.FindElementsByClassName("IEFrame");
                new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) - 100, (allFrame[0].Size.Height / 2) + 110).DoubleClick().Click().Perform();
                Thread.Sleep(2000);
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TFrmImpRep");

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
                    if (boton.Text == "OK")
                    {
                        boton.Click();
                    }
                }
                if (bandera == "0")
                {
                    errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                    if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }

                    if (pos == 0)
                    {
                        //QBE
                         QBEQry(desktopSession, bandera, QbeFilter, file);
                       // if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                    }
                }
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Error de Appium: {ex.ToString()}");
            }


            return errorMessages;
        }
        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVariables, String file)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            
            

            if (tipo == 1)
            {
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "IEFrame");
                
                PruebaCRUD.LupaDinamica(rootSession, crudVariables);
                var ElementText = rootSession.FindElementsByClassName("TEdit");
                var ElementButton = rootSession.FindElementsByClassName("TGroupButton");
                foreach (var elem in ElementButton)
                {
                    if (elem.Text == crudVariables[1]) 
                    {
                        elem.Click();
                    }
                    if (elem.Text == crudVariables[2])
                    {
                        elem.Click();
                    }

                    Thread.Sleep(500);
                }

                ElementText[0].SendKeys(crudVariables[3]);
                Screenshot("Agregar Registro", true, file, desktopSession);
            }
            


        }
    }
}
