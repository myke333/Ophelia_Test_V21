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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_8
{
    class CrudKnmcaces : FuncionesVitales
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



        static public void CrudPrincipal(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TPanel");
            //TPanel
            switch (bandera)
            {
                case 0:
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 33, 57);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1500);
                    desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                    Thread.Sleep(1500);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    desktopSession.Keyboard.SendKeys(crudPrincipal[2]);
                    Thread.Sleep(1500);
                    for (int a = 0; a < 4; a++)
                    {
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    }
                    Thread.Sleep(1500);
                    desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
                    Thread.Sleep(1500);
                    LupaDinamica(desktopSession, crudPrincipal);
                    Thread.Sleep(1500);
                    break;
                case 1:
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 513, 133);
                    desktopSession.Mouse.DoubleClick(null);
                    Thread.Sleep(1500);
                    desktopSession.Keyboard.SendKeys(crudPrincipal[4]);
                    break;
                case 2:
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 36, 26);
                    desktopSession.Mouse.Click(null);
                    break;

            }
         

        }

        static public void CrudDetalle1(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDet1)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TPanel");

            switch (bandera)
            {
                case 0:
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 100, 26);
                    desktopSession.Mouse.Click(null);
                    break;

                case 1:
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 103, 64);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1500);
                    desktopSession.Keyboard.SendKeys(crudDet1[0]);
                    Thread.Sleep(1500);
                    for (int a = 0; a < 2; a++)
                    {
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    }
                  
                    Thread.Sleep(1500);
                    desktopSession.Keyboard.SendKeys(crudDet1[1]);
                    Thread.Sleep(1500);
                    for (int a = 0; a < 2; a++)
                    {
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    }
                  
                    Thread.Sleep(1500);
                    desktopSession.Keyboard.SendKeys(crudDet1[2]);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    Thread.Sleep(1500);
                    int b = 2;
                    for (int a = 0; a < 7; a++)
                    {
                   
                        b++;
             
                        desktopSession.Keyboard.SendKeys(crudDet1[b]);
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                      

                    }
                    break;

                case 2:
                    //Por definir
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 451, 59);
                    desktopSession.Mouse.DoubleClick(null);
                    Thread.Sleep(1500);
                    desktopSession.Keyboard.SendKeys(crudDet1[10]);
                    Thread.Sleep(2000);
                    break;
            }



        }

        static public void CrudDetalle2(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDet2)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TPanel");

            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 125, 275);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudDet2[0]);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudDet2[1]);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudDet2[2]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudDet2[3]);
            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 486, 275);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudDet2[4]);
                Thread.Sleep(2000);
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

        static public void BotonoeraDetalle(WindowsDriver<WindowsElement> desktopSession, int bandera)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            switch (bandera)
            {
                case 0:
                    //agregar
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 493, 208);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(2000);
                    break;

                case 1:
                    //aceptar
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 573, 208);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(2000);

                    break;

                case 2:
                    //Editar
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 548, 208);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(2000);

                    break;
                case 3:
                    //Descartar
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 593, 208);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(2000);

                    break;
                case 4:
                    //Eliminar
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 508, 208);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(2000);
                    break;

                case 5:
                    //Dar enter
                    WindowsDriver<WindowsElement> RootSession2 = null;
                    RootSession2 = PruebaCRUD.RootSession();
                    RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(2000);
                    break;
            }



        }


    }
}
