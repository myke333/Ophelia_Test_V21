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
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_5
{ 
    class CrudKnmbasre : FuncionesVitales
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

        public static void CRUDKNmBasre(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {             
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList2 = desktopSession.FindElementsByName("Valor $");
            var ElementList3 = desktopSession.FindElementsByName("No");
            List<string> data = new List<string>() { crudVars[0] };
            if (tipo == 1)
            {
                PruebaCRUD.LupaDinamica2(desktopSession, data);
                Thread.Sleep(500);
                ElementList2[0].Click();
                Thread.Sleep(500);
                ElementList3[0].Click();
                Thread.Sleep(500);
                ElementList[6].SendKeys(crudVars[1]);
                Thread.Sleep(500);
                ElementList[5].SendKeys(crudVars[2]);
                Thread.Sleep(500);
                ElementList[27].SendKeys(crudVars[2]);
                Thread.Sleep(500);
                ElementList[37].SendKeys(crudVars[3]);
                Thread.Sleep(500);
                ElementList[34].SendKeys(crudVars[3]);
                Thread.Sleep(500);
                ElementList[36].SendKeys(crudVars[3]);
                Thread.Sleep(500);
                ElementList[28].SendKeys(crudVars[3]);
                Thread.Sleep(500);
                ElementList[26].SendKeys(crudVars[4]);
            }
            else
            {
                ElementList[26].Clear();
                ElementList[26].SendKeys(crudVars[5]);
            }
        }

       

        static public void CrudPrincipal(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudVariables)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");
            //TPanel
            if (bandera == 0)
            {
                LupaDinamica(desktopSession, crudVariables);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 710, 32);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudVariables[1]);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudVariables[2]);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 757, 350);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudVariables[3]);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates,77, 126);
                desktopSession.Mouse.DoubleClick(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);

                for (int a=0;a<29;a++)
                {
                    desktopSession.Keyboard.SendKeys(crudVariables[3]);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                }
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 576, 496);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudVariables[4]);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 674, 76);
                desktopSession.Mouse.Click(null);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1093, 28);
                desktopSession.Mouse.Click(null);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 26, 75);
                desktopSession.Mouse.Click(null);


            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 576, 496);
                desktopSession.Mouse.DoubleClick(null);
                for (int a=0;a<3;a++)
                {
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                }
                desktopSession.Keyboard.SendKeys(crudVariables[5]);

            }

        }

        public static void LupaDinamica(WindowsDriver<WindowsElement> desktopSession, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            List<Point> coordenadas = PruebaCRUD.CoordinatesFinder(desktopSession, 82, 150, 75);
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = PruebaCRUD.RootSession();
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            int contLupa = 1;
            bool IsDisplayedQbe = false;
            int indexVal = 0;
            foreach (Point coord in coordenadas)
            {
                if (indexVal == 1)
                {
                    break;
                }
                Thread.Sleep(2000);
                //Actions noteClicks = new Actions(desktopSession);
                //noteClicks.MoveToElement(parentElement).MoveByOffset(coord.X + 10, coord.Y).Click().Perform();

                desktopSession.Mouse.MouseMove(parentElement.Coordinates, coord.X + 10, coord.Y);
                desktopSession.Mouse.Click(null);
                if (contLupa == 1)
                {
                    try
                    {
                        rootSession = PruebaCRUD.RootSession();
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
                            //Thread.Sleep(2500);
                        }
                        elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                        new Actions(rootSession).MoveToElement(elements[8], 0, 0).MoveByOffset(20, 10).DoubleClick().Perform();
                        Thread.Sleep(500);
                        IsDisplayedQbe = false;

                    }
                    Thread.Sleep(1000);
                    rootSession = PruebaCRUD.RootSession();
                    rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");
                    //var campo = rootSession.FindElementsByXPath("//Window[@ClassName=\"TCaptura\"][@Name=\" LISTA DE EMPLEADOS\"]/Edit[@ClassName=\"TEdit\"]");
                    var campo = rootSession.FindElementsByClassName("TEdit");
                    ////Debugger.Launch();
                    rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                    rootSession.Mouse.Click(null);
                    Thread.Sleep(1500);
                    rootSession.Keyboard.SendKeys(data[indexVal]);
                    //Aqui se puede enviar un valor

                    var btn = rootSession.FindElementsByClassName("TBitBtn");
                    //Debugger.Launch();
                    foreach (var boton in btn)
                    {
                        if (boton.Text == "Aceptar")
                        {
                            rootSession.Mouse.MouseMove(boton.Coordinates, 10, 10);
                            rootSession.Mouse.Click(null);
                        }
                    }

                    if (indexVal == 0)
                    {
                        Thread.Sleep(20000);
                    }
                    else
                    {
                        Thread.Sleep(2000);
                    }
                    indexVal++;
                }
            }

        }


    }
}
