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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_2
{
    class CrudKnmplent:FuncionesVitales
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

        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string file)
        {
            if(tipo==1)
            {
                LupaDinamica(desktopSession, data, tipo);
                newFunctions_4.ScreenshotNuevo("Agregar Registro", true, file);
            }
            else
            {
                LupaDinamica(desktopSession, data, tipo);
                newFunctions_4.ScreenshotNuevo("Editar Registro", true, file);
            }
        }

        public static void CrudDetalle1(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, 22, 28);
            desktopSession.Mouse.Click(null);
            if (tipo == 1)
            {
                for(int i=1;i<=3;i++)
                {
                    if(i==3)
                    {
                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, 533, 28);
                        desktopSession.Mouse.Click(null);
                        Thread.Sleep(1000);
                        desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, 535, 29);
                        desktopSession.Mouse.Click(null);
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);

                    }
                    else
                    {
                        desktopSession.Keyboard.SendKeys(data[i - 1]);
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    }
                    
                }
            }
            else
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(data[data.Count - 1]);
            }
        }

        public static void CrudDetalle2(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 29, 29);
            desktopSession.Mouse.Click(null);
            if (tipo == 1)
            {
                for (int i = 1; i <= 6; i++)
                {
                    if (i == 4)
                    {
                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 626, 28);
                        desktopSession.Mouse.Click(null);
                        
                        Thread.Sleep(1000);
                       
                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 627, 29);
                        desktopSession.Mouse.Click(null);
                        
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);

                    }
                    else if (i == 6)
                    {
                        desktopSession.Keyboard.SendKeys(data[i - 1]);
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    }
                    else if(i==7)
                    {
                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 1066, 28);
                        desktopSession.Mouse.Click(null);
                        
                        Thread.Sleep(1000);

                        desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 1067, 29);
                        desktopSession.Mouse.Click(null);

                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    }
                    else
                    {
                        desktopSession.Keyboard.SendKeys(data[i - 1]);
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    }
                }
            }
            else
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(data[data.Count - 1]);
            }
        }

        public static void LupaDinamica(WindowsDriver<WindowsElement> desktopSession, List<string> data, int tipo)
        {
            List<string> errorMessages = new List<string>();
            List<Point> coordenadas = PruebaCRUD.CoordinatesFinder(desktopSession, 82, 150, 75);
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            int contLupa = 1;
            bool IsDisplayedQbe = false;
            int indexVal = 0;
            foreach (Point coord in coordenadas)
            {
                Thread.Sleep(500);
                if (tipo == 1)
                {
                    if (contLupa == 1 )
                    {
                        desktopSession.Mouse.MouseMove(parentElement.Coordinates, coord.X + 10, coord.Y);
                        desktopSession.Mouse.Click(null);

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
                    contLupa++;
                }
                else
                {
                    if (contLupa == 1)
                    {
                        desktopSession.Mouse.MouseMove(parentElement.Coordinates, coord.X + 10, coord.Y);
                        desktopSession.Mouse.Click(null);

                        rootSession = RootSession();
                        rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");
                        //var campo = rootSession.FindElementsByXPath("//Window[@ClassName=\"TCaptura\"][@Name=\" LISTA DE EMPLEADOS\"]/Edit[@ClassName=\"TEdit\"]");
                        var campo = rootSession.FindElementsByClassName("TEdit");
                        ////Debugger.Launch();
                        rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                        rootSession.Mouse.Click(null);
                        rootSession.Keyboard.SendKeys(data[1]);
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
                        contLupa++;
                    }
                }
            }
        }

        public static void EliminarRegistro(WindowsDriver<WindowsElement> desktopSession, WindowsElement ElementList, Dictionary<string, Point> botonesNavegador, string file)
        {
            WindowsDriver<WindowsElement> rootSession = null;
            desktopSession.Mouse.MouseMove(ElementList.Coordinates, botonesNavegador["Borrar"].X, botonesNavegador["Borrar"].Y);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Borrar Registro", true, file);
            rootSession= RootSession();
            rootSession = ReloadSession(rootSession, "TMessageForm");
            var ventana = rootSession.FindElementsByName("Confirm");
            ventana[0].Click();
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Registro Borrado", true, file);
            Thread.Sleep(2000);
        }
        
    }
}
