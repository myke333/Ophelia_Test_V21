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
    class CrudKbicargo:FuncionesVitales
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
                ElementList[11].SendKeys(variables[0]);
                ElementList[12].SendKeys(variables[1]);
                ElementList2[0].SendKeys(variables[2]);
                ElementList[7].SendKeys(variables[3]);
                ElementList[8].SendKeys(variables[5]);
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
                newFunctions_4.ScreenshotNuevo("Registro Agregado", true, file);

            }
            else
            {
                newFunctions_4.ScreenshotNuevo("Editar Registro", true, file);
                ElementList2[0].Click();
                ElementList2[0].Clear();
                ElementList2[0].SendKeys(variables[7]);
                newFunctions_4.ScreenshotNuevo("Registro Editado", true, file);
            }
        }


        static public void CrudDetalle1y2(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDet1, string file)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");
            //TPanel

            switch (bandera)
            {
                case 0:
                    //ventana mision
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 116, 10);
                    desktopSession.Mouse.Click(null);
                    newFunctions_4.ScreenshotNuevo("Ventana mision ", true, file);
                
                    break;
                case 1:
                    //ventana requisitos
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 221, 10);
                    desktopSession.Mouse.Click(null);
                    newFunctions_4.ScreenshotNuevo("Ventana requisitos ", true, file);
                    break;
                case 2:
                    //desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 29, 18);
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 122, 65);
                    desktopSession.Mouse.Click(null);
                    desktopSession.Keyboard.SendKeys(crudDet1[0]);
                    break;
                case 3:
                    //editar
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 122, 65);
                    desktopSession.Mouse.DoubleClick(null);
                    desktopSession.Keyboard.SendKeys(crudDet1[1]);
                    break;
            }

        }

        static public void CrudDetalle3(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDet2, string file)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");
            //TPanel
            var btnTDVI2 = desktopSession.FindElementsByClassName("TDBGrid");
            switch (bandera)
            {
                case 0:
                    //ventana mision
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 350, 10);
                    desktopSession.Mouse.Click(null);
                    newFunctions_4.ScreenshotNuevo("Ventana Analisis ", true, file);
                    break;
                case 1:
                    desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 53, 31);
                    desktopSession.Mouse.Click(null);
                    desktopSession.Keyboard.SendKeys(crudDet2[0]);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    desktopSession.Keyboard.SendKeys(crudDet2[0]);
                    for (int a=0;a<12;a++)
                    {
                        desktopSession.Keyboard.SendKeys(crudDet2[1]);
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    }

                    break;
                case 2:
                    desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 1147, 31);
                    desktopSession.Mouse.Click(null);
                    desktopSession.Keyboard.SendKeys(crudDet2[2]);
                    break;

            }

        }



        static public void CrudDetalle(WindowsDriver<WindowsElement> desktopSession, int bandera, string file)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");
            //TPanel
            var btnTDVI2 = desktopSession.FindElementsByClassName("TDBGrid");
            switch (bandera)
            {
              
                case 0:
                    //agregar edicion
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 10, 166);
                    desktopSession.Mouse.Click(null);
                    newFunctions_4.ScreenshotNuevo("Agregar registro ventana detalle", true, file);

                    break;
                case 1:
                    //Aceptaredicion
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 34, 154);
                    desktopSession.Mouse.Click(null);
                    newFunctions_4.ScreenshotNuevo("Aceptar edicion", true, file);

                    break;

                case 2:
                    //rechazaredicion
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 67, 154);
                    desktopSession.Mouse.Click(null);
                    newFunctions_4.ScreenshotNuevo("Aceptar edicion", true, file);

                    break;

            }

        }

        static public void Enter(WindowsDriver<WindowsElement> desktopSession)
        {
            WindowsDriver<WindowsElement> RootSession2 = null;
            RootSession2 = PruebaCRUD.RootSession();
            RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2000);

        }


    }
}
