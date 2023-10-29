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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBi
{
    class CrudKbihcarg:FuncionesVitales
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

        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> variables, string file)
        {
            List<string> errorMessages = new List<string>();
            List<Point> coordenadas = PruebaCRUD.CoordinatesFinder(desktopSession, 82, 150, 75);
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            int contLupa = 0;

            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            int cont = 0;
            var ElementList2 = desktopSession.FindElementsByClassName("TDBMemo");
            //foreach (var elem in ElementList)
            //{
            //    elem.SendKeys(cont.ToString());
            //    cont++;
            //    Thread.Sleep(2000);
            //}
            if (tipo == 1)
            {
                newFunctions_4.ScreenshotNuevo("Agregar Registro", true, file);
                ElementList[21].SendKeys(variables[0]);
                ElementList[22].SendKeys(variables[1]);
                ElementList[23].SendKeys(variables[2]);
                ElementList[8].SendKeys(variables[3]);
                ElementList[7].SendKeys(variables[5]);
                ElementList[5].SendKeys(variables[6]);
                foreach (Point coord in coordenadas)
                {
                    Thread.Sleep(2000);
                    //Actions noteClicks = new Actions(desktopSession);
                    //noteClicks.MoveToElement(parentElement).MoveByOffset(coord.X + 10, coord.Y).Click().Perform();

                    desktopSession.Mouse.MouseMove(parentElement.Coordinates, coord.X + 10, coord.Y);
                    desktopSession.Mouse.Click(null);

                    rootSession = RootSession();
                    rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");
                    //var campo = rootSession.FindElementsByXPath("//Window[@ClassName=\"TCaptura\"][@Name=\" LISTA DE EMPLEADOS\"]/Edit[@ClassName=\"TEdit\"]");
                    var campo = rootSession.FindElementsByClassName("TEdit");
                    ////Debugger.Launch();
                    rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                    rootSession.Mouse.Click(null);
                    rootSession.Keyboard.SendKeys(variables[4]);
                    contLupa++;

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
                newFunctions_4.ScreenshotNuevo("Registro Agregaedo", true, file);

            }
            else
            {
                newFunctions_4.ScreenshotNuevo("Editar Registro", true, file);
                ElementList[23].Click();
                ElementList[23].Clear();
                ElementList[23].SendKeys(variables[7]);
                newFunctions_4.ScreenshotNuevo("Registro Editado", true, file);
            }
        }
    }
}
