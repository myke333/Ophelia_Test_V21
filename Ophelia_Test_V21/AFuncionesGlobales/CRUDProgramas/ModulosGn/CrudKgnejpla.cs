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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBi
{
    class CrudKgnejpla : FuncionesVitales
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

            var btnTDVI = desktopSession.FindElementsByClassName("TDBEdit");
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");
            //var btnTDVI = desktopSession.FindElementsByClassName("TPanel");
            //Debugger.Launch();


            if (bandera == 0)
            {
                btnTDVI[9].SendKeys(crudPrincipal[0]);
                Thread.Sleep(1500);
                btnTDVI[7].SendKeys(crudPrincipal[1]);
                Thread.Sleep(1500);
                btnTDVI[8].SendKeys(crudPrincipal[0]);
                Thread.Sleep(1500);
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                //Thread.Sleep(1500);



                //desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 101, 27);
                //desktopSession.Mouse.Click(null);
                //desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                //Thread.Sleep(1500);

            }
            else
            {
                btnTDVI[9].Clear();
                btnTDVI[9].SendKeys(crudPrincipal[2]);
                Thread.Sleep(1500);
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                //Thread.Sleep(1500);

                //desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 101, 27);
                //desktopSession.Mouse.Click(null);
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                //desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
                //Thread.Sleep(1500);

            }
        }

        static public void ClickButtonAceptarEnter(WindowsDriver<WindowsElement> desktopSession, int bandera, string file)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TBitBtn");
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");
            //var btnTDVI = desktopSession.FindElementsByClassName("TPanel");
            //Debugger.Launch();


            if (bandera == 0)
            {
                    var ElementList = desktopSession.FindElementsByName("Aceptar");
                    ElementList[0].Click();
                

            }
            else
            {
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1500);
                WindowsDriver<WindowsElement> RootSession = null;
                RootSession = PruebaCRUD.RootSession();
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                //Thread.Sleep(1500);

                //desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 101, 27);
                //desktopSession.Mouse.Click(null);
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                //desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
                //Thread.Sleep(1500);

            }
        }


        static public void Click(WindowsDriver<WindowsElement> desktopSession, int bandera)
        {

            //var btnTDVI2 = desktopSession.FindElementsByClassName("TDBEdit");
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");
            //var btnTDVI = desktopSession.FindElementsByClassName("TPanel");
            //Debugger.Launch();

            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 84, 50);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 84, 68);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 148, 83);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                //WindowsDriver<WindowsElement> RootSession = null;
                //RootSession = PruebaCRUD.RootSession();
                //RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 22, 570);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                //WindowsDriver<WindowsElement> RootSession = null;
                //RootSession = PruebaCRUD.RootSession();
                //RootSession.Keyboard.SendKeys("ReportePDF1_Preliminar_08022022_161632");
                //Thread.Sleep(1000);
                //RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

            }
        }

        static public void generarExcel(WindowsDriver<WindowsElement> Session, string file, string ruta)
        {
            List<string> errorMessages = new List<string>();
            
                            Session = RootSession();
                           
                                //var saveExcel = Session.FindElementsByName("Maximizar");
                                //var cantidad = saveExcel.Count;
                                //if (cantidad == 2)
                                //{
                                //    saveExcel[1].Click();
                                //}
                                //else
                                //{
                                //    saveExcel[0].Click();
                                //}
                                //Cambio Jose
                                Session.FindElementByName("Pestaña Archivo").Click();
                                //Fin Cambio Jose

                                //Session.FindElementByName("Guardar").Click();
                                Session.FindElementByName("Guardar como").Click();
                                Session.FindElementByName("Examinar").Click();
                                Thread.Sleep(1000);
                                Session.Keyboard.SendKeys(ruta);
                                Session.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                Thread.Sleep(1000);
                                newFunctions_4.ScreenshotNuevo("Reporte Excel Guardado", true, file);
                                LimpiarProcesos();

                            
        }

        static public void Clickk(WindowsDriver<WindowsElement> desktopSession, int bandera)
        {

            //var btnTDVI2 = desktopSession.FindElementsByClassName("TDBEdit");
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");
            //var btnTDVI2 = desktopSession.FindElementsByClassName("TTabSheet");
            //Debugger.Launch();

            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 468, 190);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                //desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 91, -10);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1000);
                //WindowsDriver<WindowsElement> RootSession = null;
                //RootSession = PruebaCRUD.RootSession();
                //RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                //Thread.Sleep(1000);
                //desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 586, 170);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1000);
                //RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                //Thread.Sleep(1000);

            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 22, 570);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                //WindowsDriver<WindowsElement> RootSession = null;
                //RootSession = PruebaCRUD.RootSession();
                //RootSession.Keyboard.SendKeys("ReportePDF1_Preliminar_08022022_161632");
                //Thread.Sleep(1000);
                //RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

            }
        }



    }
}
