using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Drawing;
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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_9
{
    class CrudKnmvicap : FuncionesVitales
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

        static public void AgregarCrud(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal)
        {
            //Debugger.Launch();
            var btnTDVI = desktopSession.FindElementsByClassName("TDBEdit");
            var btnTDVI1 = desktopSession.FindElementsByClassName("TPanel");
            Thread.Sleep(5000);
            /*for (int i = 0; i < btnTDVI.Count; i++)
            {
                btnTDVI[i].SendKeys(i.ToString());
                Thread.Sleep(1800);
            }*/
            /*if(bandera == 0)
            {
                LupaDinamica(desktopSession, crudPrincipal);
            }
            else
            {
                btnTDVI[7].SendKeys(crudPrincipal[12]);
            }*/
            if (bandera == 0)
            {
                btnTDVI[7].SendKeys(crudPrincipal[12]);
                Thread.Sleep(1800);
                btnTDVI[6].SendKeys(crudPrincipal[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(2500);
                desktopSession.Mouse.MouseMove(btnTDVI1[0].Coordinates, 86, 232);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                desktopSession.Keyboard.SendKeys(crudPrincipal[4]);
                //btnTDVI[19].SendKeys(crudPrincipal[4]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                //Thread.Sleep(1500);

            }
            else if (bandera== 1)
            {
                btnTDVI[19].Clear();
                btnTDVI[19].SendKeys(crudPrincipal[11]);
                Thread.Sleep(3000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(5000);
            }
            else
            {
                btnTDVI[7].SendKeys(crudPrincipal[12]);
                Thread.Sleep(5000);
            }

        }

        static public void ClickButton(WindowsDriver<WindowsElement> desktopSession, int bandera)
        {
            //var btnTDVI = desktopSession.FindElementsByClassName("TDBGrid");

            if (bandera == 0)
            {
                desktopSession.FindElementByName("Información Básica").Click();
            }
            else if (bandera == 1)
            {
                desktopSession.FindElementByName("Solicitud").Click();
            }
            else
            {
                desktopSession.FindElementByName("Descripción").Click();
            }

        }

        static public void AgregarCrud1(WindowsDriver<WindowsElement> desktopSession, List<string> crudPrincipal)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TTabSheet");
                                               
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 115, 38);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[5]);
                Thread.Sleep(4000);

                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 115, 80);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[6]);
                Thread.Sleep(4000);

                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 55, 80);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[7]);
                Thread.Sleep(4000);

                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 315, 35);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(3000);
                WindowsDriver<WindowsElement> rootSesison = null;
                rootSesison = PruebaCRUD.RootSession();
                Thread.Sleep(3000);
                rootSesison.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(3000);

                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 315, 80);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(3000);
                rootSesison = PruebaCRUD.RootSession();
                Thread.Sleep(3000);
                rootSesison.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(3000); 

                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 255, 172);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[8]);
                Thread.Sleep(3000);

                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 520, 216);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[9]);
                Thread.Sleep(3000);
        }

        static public void AgregarCrud2(WindowsDriver<WindowsElement> desktopSession, List<string> crudPrincipal)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("(TDBMemo)");

            //desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 63, 26);
            //desktopSession.Mouse.Click(null);
            desktopSession.Keyboard.SendKeys(crudPrincipal[10]);                       

        }


        static public void Enter(WindowsDriver<WindowsElement> desktopSession)
        {


            //var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");
            var btnTDVI = desktopSession.FindElementsByClassName("TPanel");


            //Coordenadas boton qbe2
            //string filterqb = "13";


            Thread.Sleep(1000);
            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 755, 266);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);

       
            //Thread.Sleep(2000);
            //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

            //WindowsDriver<WindowsElement> RootSession = null;
            //RootSession = PruebaCRUD.RootSession();
            //RootSession = ReloadSession(RootSession, "TFrmInfoP50");
            //Thread.Sleep(1000);
            //RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

            
        }

        static public void Preliminar(WindowsDriver<WindowsElement> desktopSession)
        {


            //var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");
            var btnTDVI = desktopSession.FindElementsByClassName("TPanel");


            //Coordenadas boton qbe2
            //string filterqb = "13";
            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 128, 13);
            desktopSession.Mouse.Click(null);

            Thread.Sleep(2000);
            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 706, 460);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);


            //Thread.Sleep(2000);
            //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

            //WindowsDriver<WindowsElement> RootSession = null;
            //RootSession = PruebaCRUD.RootSession();
            //RootSession = ReloadSession(RootSession, "TFrmInfoP50");
            //Thread.Sleep(1000);
            //RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);


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
                if (indexVal == 5)
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
