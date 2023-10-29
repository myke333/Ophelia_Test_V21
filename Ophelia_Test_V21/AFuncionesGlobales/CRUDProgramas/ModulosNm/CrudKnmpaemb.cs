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
    class CrudKnmpaemb:FuncionesVitales
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

        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data,string file)
        {
            if (data[0]=="SQL")
            {
                if (tipo == 1)
                {
                    newFunctions_4.ScreenshotNuevo("Agregar Registro", true, file);
                    LupaDinamica(desktopSession, data, tipo);
                }
                else
                {
                    newFunctions_4.ScreenshotNuevo("Editar Registro", true, file);
                    LupaDinamica(desktopSession, data, tipo);
                }
            }
            else
            {
                if (tipo == 1)
                {
                    var ElementList = desktopSession.FindElementsByClassName("TDBCheckBox");
                    newFunctions_4.ScreenshotNuevo("Agregar Registro", true, file);
                    LupaDinamica2(desktopSession, data, tipo);
                    foreach(var elem in ElementList)
                    {
                        if(elem.Text=="Valida Detalles?" || elem.Text=="Genera IVA?")
                        {
                            elem.Click();
                            if(elem.Text == "Genera IVA?")
                            {
                                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                                desktopSession.Keyboard.SendKeys(data[data.Count-1]);
                            }
                        }
                    }
                }
                else
                {
                    newFunctions_4.ScreenshotNuevo("Editar Registro", true, file);
                    LupaDinamica2(desktopSession, data, tipo);
                }
            }
            
            
        }
        public static void LupaDinamica(WindowsDriver<WindowsElement> desktopSession, List<string> data,int tipo)
        {
            //Debugger.Launch();
            List<string> errorMessages = new List<string>();
            List<Point> coordenadas = PruebaCRUD.CoordinatesFinder(desktopSession, 82, 150, 75);
            Debug.WriteLine($"Numero de Lupas {coordenadas.Count}");
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            int contLupa = 1;
            bool IsDisplayedQbe = false;
            int indexVal = 1;
            int contqbe = 0;
            if(tipo==1)
            { 
                foreach (Point coord in coordenadas)
                {
                    Thread.Sleep(2000);

                    desktopSession.Mouse.MouseMove(parentElement.Coordinates, coord.X + 10, coord.Y);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(3000);
                    if (contLupa == 1)
                    {
                        rootSession = RootSession();
                        rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");
                        //var campo = rootSession.FindElementsByXPath("//Window[@ClassName=\"TCaptura\"][@Name=\" LISTA DE EMPLEADOS\"]/Edit[@ClassName=\"TEdit\"]");
                        var campo = rootSession.FindElementsByClassName("TEdit");
                        Thread.Sleep(1000);
                        ////Debugger.Launch();
                        rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                        rootSession.Mouse.Click(null);
                        Thread.Sleep(1000);
                        rootSession.Keyboard.SendKeys(data[indexVal]);
                        Thread.Sleep(1000);
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
            else
            {
                desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordenadas[9].X + 10, coordenadas[9].Y);
                desktopSession.Mouse.Click(null);
                rootSession = RootSession();
                rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");
                //var campo = rootSession.FindElementsByXPath("//Window[@ClassName=\"TCaptura\"][@Name=\" LISTA DE EMPLEADOS\"]/Edit[@ClassName=\"TEdit\"]");
                var campo = rootSession.FindElementsByClassName("TEdit");
                ////Debugger.Launch();
                rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                rootSession.Mouse.Click(null);
                rootSession.Keyboard.SendKeys(data[10]);
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
            }
        }

        public static void LupaDinamica2(WindowsDriver<WindowsElement> desktopSession, List<string> data, int tipo)
        {
            List<string> errorMessages = new List<string>();
            List<Point> coordenadas = PruebaCRUD.CoordinatesFinder(desktopSession, 82, 150, 75);
            Debug.WriteLine($"Numero de Lupas {coordenadas.Count}");
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            int contLupa = 1;
            bool IsDisplayedQbe = false;
            int indexVal = 1;
            int cont=1;
            if (tipo == 1)
            {
                foreach (Point coord in coordenadas)
                {
                    if(cont== 1 || cont ==2 || cont ==3 || cont==5 || cont == 7)
                    {
                        Thread.Sleep(1000);
                        
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
                    cont++;
                }
            }
            else
            {
                desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordenadas[9].X + 10, coordenadas[9].Y);
                desktopSession.Mouse.Click(null);
                rootSession = RootSession();
                rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");
                //var campo = rootSession.FindElementsByXPath("//Window[@ClassName=\"TCaptura\"][@Name=\" LISTA DE EMPLEADOS\"]/Edit[@ClassName=\"TEdit\"]");
                var campo = rootSession.FindElementsByClassName("TEdit");
                ////Debugger.Launch();
                rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                rootSession.Mouse.Click(null);
                rootSession.Keyboard.SendKeys(data[10]);
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
            }
        }
    }
}
