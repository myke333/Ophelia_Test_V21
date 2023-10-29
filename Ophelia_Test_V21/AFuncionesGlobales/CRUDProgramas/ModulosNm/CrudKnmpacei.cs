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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm
{
    class CrudKnmpacei : FuncionesVitales
    {

        static public WindowsDriver<WindowsElement> RootSession()
        {

            var appCapabilities = new AppiumOptions();
            appCapabilities.AddAdditionalCapability("app", "Root");
            var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
            return ApplicationSession;
        }

        static public WindowsDriver<WindowsElement> RootSessionNew()
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
        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVariables, String file)
        {
            List<string> errorMessages = new List<string>();
            List<int> lupasOff = new List<int>() { 0 };
            WindowsDriver<WindowsElement> rootSession = null;
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementPage = desktopSession.FindElementsByClassName("TPageControl");
            var ElementBox = desktopSession.FindElementsByClassName("TKsCmbCodigo");
            var ElementCombo = desktopSession.FindElementsByClassName("TDBComboBox");

            

            if (tipo == 1)
            {
                ElementList[9].SendKeys(crudVariables[0]);
                ElementList[7].SendKeys(crudVariables[1]);
                ElementList[8].SendKeys(crudVariables[2]);
                ElementList[17].SendKeys(crudVariables[3]);
                ElementList[15].SendKeys(crudVariables[4]);
                ElementList[13].SendKeys(crudVariables[5]);
                ElementList[11].SendKeys(crudVariables[6]);
                ElementList[6].SendKeys(crudVariables[7]);
                ElementCombo[0].SendKeys(crudVariables[26]);
                ElementList[5].SendKeys(crudVariables[8]);
                ElementList[4].SendKeys(crudVariables[9]);
                ElementList[3].SendKeys(crudVariables[10]);
                ElementList[2].SendKeys(crudVariables[11]);
                

                desktopSession.Mouse.MouseMove(ElementPage[0].Coordinates, 22, 7);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(6000);

                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                rootSession = RootSession();

                rootSession = ReloadSession(rootSession, "IEFrame");
                ElementList = desktopSession.FindElementsByClassName("TDBEdit");
                var ElementButton = desktopSession.FindElementsByClassName("TGroupButton");
                var ElementMemo = desktopSession.FindElementsByClassName("TDBMemo");
                
               

                ElementList[14].SendKeys(crudVariables[3]);
                ElementList[13].SendKeys(crudVariables[4]);
                ElementList[12].SendKeys(crudVariables[5]);
                ElementList[8].SendKeys(crudVariables[6]);
                ElementList[2].SendKeys(crudVariables[12]);
                ElementList[6].SendKeys(crudVariables[13]);
                ElementList[4].SendKeys(crudVariables[14]);
                ElementList[5].SendKeys(crudVariables[15]);
                ElementList[3].SendKeys(crudVariables[16]);

                foreach (var elem in ElementButton)
                {
                    if (elem.Text == crudVariables[24])
                    {
                        elem.Click();
                    }
                    Thread.Sleep(500);
                }
                ElementMemo[0].SendKeys(crudVariables[25]);
                ElementBox[0].Click();

                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                Thread.Sleep(6000);
                rootSession = RootSession();

                rootSession = ReloadSession(rootSession, "IEFrame");
                LupaDinamica(desktopSession, crudVariables, 17);

               // newFunctions_4.findText(desktopSession, crudVariables, "TDBEdit");

            }
            else
            {
                desktopSession.Mouse.MouseMove(ElementPage[0].Coordinates, 22, 7);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(6000);
                rootSession = RootSession();

                rootSession = ReloadSession(rootSession, "IEFrame");
                ElementList = desktopSession.FindElementsByClassName("TDBEdit");

                ElementList[2].Clear();
                ElementList[2].SendKeys(crudVariables[27]);
            }


        }

        public static void LupaDinamica(WindowsDriver<WindowsElement> desktopSession, List<string> data, int val)
        {
            List<string> errorMessages = new List<string>();
            List<Point> coordenadas = PruebaCRUD.CoordinatesFinder(desktopSession, 82, 150, 75);
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            int contLupa = 1;
            bool IsDisplayedQbe = false;
            int indexVal = val;
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
                        rootSession = PruebaCRUD.ReloadSession(rootSession, "TKQBE");
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
                    rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");
                    //var campo = rootSession.FindElementsByXPath("//Window[@ClassName=\"TCaptura\"][@Name=\" LISTA DE EMPLEADOS\"]/Edit[@ClassName=\"TEdit\"]");
                    var campo = rootSession.FindElementsByClassName("TEdit");
                    ////Debugger.Launch();
                    rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                    rootSession.Mouse.Click(null);
                    rootSession.Keyboard.SendKeys(data[indexVal]);
                    //Aqui se puede enviar un valor

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
    }
}
