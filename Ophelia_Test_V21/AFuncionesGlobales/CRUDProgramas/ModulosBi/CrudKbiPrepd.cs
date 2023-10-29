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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBi
{
    class CrudKbiPrepd :FuncionesVitales
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

        public static void LupaDinamica(WindowsDriver<WindowsElement> desktopSession, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            List<Point> coordenadas = PruebaCRUD.CoordinatesFinder(desktopSession, 82, 150, 75);
            Debug.WriteLine($"Numero de Lupas {coordenadas.Count}");
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            int contLupa = 1;
            bool IsDisplayedQbe = false;
            int indexVal = 0;
            int contqbe = 0;
            if (contLupa == 1)
            {
                foreach (Point coord in coordenadas)
                {
                    Thread.Sleep(2000);

                    desktopSession.Mouse.MouseMove(parentElement.Coordinates, coord.X + 10, coord.Y);
                    desktopSession.Mouse.Click(null);
                    if (contLupa == 1)
                    {
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



        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVariables, String file)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;


            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            var ElementText = desktopSession.FindElementsByClassName("TDBMemo");
            var ElementButton = desktopSession.FindElementsByClassName("TGroupButton");
            var ElementList3 = desktopSession.FindElementsByClassName("TDBGridInplaceEdit");  

            if (tipo == 1)
            {
                LupaDinamica(desktopSession, crudVariables);
                              
            }
            else if(tipo==2)
            {

                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates,27,27);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(6000);

                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates,140, 28);
                desktopSession.Mouse.Click(null);

                rootSession = RootSession();
               
                var ventana = rootSession.FindElementsByClassName("TDBMemo");
                ventana[0].SendKeys(crudVariables[1]);
                var btn= rootSession.FindElementsByClassName("TBitBtn");
                foreach(var boton in btn)
                {
                    if(boton.Text=="Aceptar")
                    {
                        boton.Click();
                    }
                }

            }
            else if(tipo==3)
            {
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 27, 27);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(6000);

                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 140, 28);
                desktopSession.Mouse.Click(null);

                rootSession = RootSession();
                var ventana = rootSession.FindElementsByClassName("TDBMemo");
                ventana[0].Clear();
                ventana[0].SendKeys(crudVariables[2]);
                var btn = rootSession.FindElementsByClassName("TBitBtn");
                foreach (var boton in btn) 
                {
                    if (boton.Text == "Aceptar")
                    {
                        boton.Click();
                    }
                }

            }
            else 
            {
                foreach (var elem in ElementButton)
                {
                    if (elem.Text == crudVariables[3])
                    {
                        elem.Click();
                    }
                    Thread.Sleep(500);
                }
            }
        }
       
        public static void cerrarBorrar(WindowsDriver<WindowsElement> desktopSession)
        {
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            rootSession = ReloadSession(rootSession, "IEFrame");
            var allFrame = rootSession.FindElementsByClassName("IEFrame");
            new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset(allFrame[0].Size.Width / 2, allFrame[0].Size.Height / 2).DoubleClick().Click().Perform();
            rootSession = RootSession();
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
        }
    }
}
