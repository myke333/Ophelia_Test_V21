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
    class CrudKnmctrel:FuncionesVitales
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
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList2 = desktopSession.FindElementsByClassName("TGroupBox");

            if (tipo == 1)
            {
                newFunctions_4.ScreenshotNuevo("Agregar Registro", true, file);
                foreach (var elem in ElementList2)
                {
                    if (elem.Text == "Formatos")
                    {
                        desktopSession.Mouse.MouseMove(elem.Coordinates, 105, 46);
                        desktopSession.Mouse.Click(null);
                        Thread.Sleep(200);
                        rootSession = RootSession();
                        rootSession = ReloadSession(rootSession, "TFrmFmtFecha");
                        var btn1 = rootSession.FindElementsByClassName("TBitBtn");
                        foreach (var btn in btn1)
                        {
                            if (btn.Text == "OK")
                            {
                                btn.Click();
                            }
                        }
                        desktopSession.Mouse.MouseMove(elem.Coordinates, 194, 54);
                        desktopSession.Mouse.Click(null);
                        rootSession = RootSession();
                        rootSession = ReloadSession(rootSession, "TFrmFmtHora");
                        var btn2 = rootSession.FindElementsByClassName("TBitBtn");
                        foreach (var btn in btn2)
                        {
                            if (btn.Text == "OK")
                            {
                                btn.Click();
                            }
                        }
                    }
                }
                LupaDinamica(desktopSession, data);
                ElementList[5].SendKeys(data[1]);
                ElementList[4].SendKeys(data[2]);
                ElementList[7].SendKeys(data[3]);
                ElementList[6].SendKeys(data[4]);
            }
            else
            {
                newFunctions_4.ScreenshotNuevo("Editar Registro", true, file);
                ElementList[4].Clear();
                ElementList[4].SendKeys(data[5]);
            }


        }
        public static void LupaDinamica(WindowsDriver<WindowsElement> desktopSession, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            List<Point> coordenadas = PruebaCRUD.CoordinatesFinder(desktopSession, 82, 150, 75);
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            int contLupa = 1;
            bool IsDisplayedQbe = false;
            int indexVal = 0;
            int contqbe = 0;
            foreach (Point coord in coordenadas)
            {
                Thread.Sleep(2000);
                //Actions noteClicks = new Actions(desktopSession);
                //noteClicks.MoveToElement(parentElement).MoveByOffset(coord.X + 10, coord.Y).Click().Perform();

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
}
