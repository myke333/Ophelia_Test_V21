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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSO
{
    class CrudKsocovac : FuncionesVitales
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

        static public WindowsDriver<WindowsElement> ReloadSession1(WindowsDriver<WindowsElement> session, string parametro)
        {
            try
            {
                var ApplicationWindow = session.FindElementByClassName(parametro);
                var ApplicationSessionHandle = ApplicationWindow.GetAttribute("NativeWindowHandle");
                ApplicationSessionHandle = (int.Parse(ApplicationSessionHandle)).ToString("x"); //Convert to hex


                var appCapabilities = new AppiumOptions();
                appCapabilities.AddAdditionalCapability("appTopLevelWindow", ApplicationSessionHandle);
                var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);

                string ses = ApplicationSession.PageSource;
                return ApplicationSession;
            }
            catch (Exception ex) { return null; }
        }

        static public void AgregarCrud(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal)
        {

            //var btnTDVI = desktopSession.FindElementsByClassName("TDBEdit");
            //var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");
            //var btnTDVI = desktopSession.FindElementsByClassName("TTabSheet");
            //Debugger.Launch();
            //for (int i = 0; i < btnTDVI.Count; i++)
            //{
            //    btnTDVI[i].SendKeys(i.ToString());
            //}
            var btnTDVI = desktopSession.FindElementsByClassName("TDBEdit");
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");


            if (bandera == 0)
            {
                LupaDinamica(desktopSession, crudPrincipal);
                Thread.Sleep(1500);
            }
            else
            {
                LupaDinamica(desktopSession, crudPrincipal);
            }

        }

        static public void AgregarCrudDetalle(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDetalle)
        {

            //var btnTDVI = desktopSession.FindElementsByClassName("TDBEdit");
            //var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");
            //var btnTDVI = desktopSession.FindElementsByClassName("TTabSheet");
            //Debugger.Launch();
            //for (int i = 0; i < btnTDVI.Count; i++)
            //{
            //    btnTDVI[i].SendKeys(i.ToString());
            //}
            var btnTDVI = desktopSession.FindElementsByClassName("TDBEdit");
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");


            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 120, 160);
                desktopSession.Mouse.Click(null);

                Thread.Sleep(1500);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 120, 164);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1500);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 349, 174);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 349, 277);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);


                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 754, 169);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudDetalle[0]);

                Thread.Sleep(1500);

                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudDetalle[3]);

                Thread.Sleep(1500);

                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudDetalle[1]);

                Thread.Sleep(1500);

                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudDetalle[3]);

                Thread.Sleep(1500);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 918, 295);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudDetalle[4]);
                Thread.Sleep(1500);
            }
            else
            {


                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 349, 377);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1500);


                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 754, 469);
                desktopSession.Mouse.DoubleClick(null);
                desktopSession.Keyboard.SendKeys(crudDetalle[0]);


                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 754, 169);
                desktopSession.Mouse.DoubleClick(null);
                desktopSession.Keyboard.SendKeys(crudDetalle[1]);

                Thread.Sleep(1500);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 754, 235);
                desktopSession.Mouse.DoubleClick(null);
                desktopSession.Keyboard.SendKeys(crudDetalle[2]);
                Thread.Sleep(1500);


                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 754, 295);
                desktopSession.Mouse.DoubleClick(null);
                desktopSession.Keyboard.SendKeys(crudDetalle[5]);
                Thread.Sleep(1500);
            }
        }


        static public void Enter(WindowsDriver<WindowsElement> desktopSession)
        {
            WindowsDriver<WindowsElement> RootSession2 = null;
            RootSession2 = PruebaCRUD.RootSession();
            RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2000);

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
                if (indexVal == 3)
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

        static public void Seguir(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDetalle)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TDBEdit");
            var btnTDVI2 = desktopSession.FindElementsByClassName("TKactusDBgrid");

            desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 530, 323);
            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2000);
            desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 600, 24);
            desktopSession.Mouse.DoubleClick(null);
        }

        static public void reportePreliminar(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDetalle)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TGridPanel");
            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 130, 17);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);

            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession = ReloadSession(RootSession, "TFrmMenu");

            var btnTDVI2 = RootSession.FindElementsByClassName("TPanel");

            if (bandera == 0)
            {
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
            else
            {
                RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 15, 166);
                RootSession.Mouse.Click(null);
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
           
         
        }



    }
}
