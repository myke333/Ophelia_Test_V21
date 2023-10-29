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
    class CrudKnmvamar : FuncionesVitales
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



        static public void CrudPrincipal(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");
            //TPanel
            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 165, 48);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                Thread.Sleep(1500);
            }
            else
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Space);
                Thread.Sleep(1500);
                //desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 432, 23);
                //desktopSession.Mouse.Click(null);
            }


        }

        static public void Botones(WindowsDriver<WindowsElement> desktopSession, int bandera)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");
            //TPanel
            switch (bandera)
            {
                case 0:
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 513, 69);
                    desktopSession.Mouse.Click(null);
                    break;
                case 1:
                    //excel, por si no sirve
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 928, 282);
                    desktopSession.Mouse.Click(null);
                    break;
                case 2:
                    //Boton preliminar:
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 657, 282);
                    desktopSession.Mouse.Click(null);
                    break;


            }

        }

        static public void ExpExel(WindowsDriver<WindowsElement> Session, string ruta, string file)
        {


            //Abre ventana excel
            Session = RootSession();
            Session = ReloadSession1(Session, "XLMAIN");

            if (Session != null)

            {
                int count = 0;
                while (count < 240)
                {
                    try
                    {
                        //Session = RootSession();
                        //Session = ReloadSession1(Session, "XLMAIN");
                        Thread.Sleep(500);
                        var saveExcel1 = Session.FindElementsByName("Maximizar");
                        if (saveExcel1.Count > 0)
                        {
                            break;
                        }
                    }
                    catch
                    {
                        Thread.Sleep(1000);
                        count++;
                    }
                }
                var saveExcel = Session.FindElementsByName("Maximizar");
                var cantidad = saveExcel.Count;
                if (cantidad == 2)
                {
                    saveExcel[1].Click();
                }
                else
                {
                    saveExcel[0].Click();
                }
                //Cambio Jose
                Session.FindElementByName("Pestaña Archivo").Click();
                //Fin Cambio Jose

                //Session.FindElementByName("Guardar").Click();
                Session.FindElementByName("Guardar").Click();
                Session.FindElementByName("Examinar").Click();
                Thread.Sleep(1000);
                Session.Keyboard.SendKeys(ruta);
                Session.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);
                newFunctions_4.ScreenshotNuevo("Reporte Excel Guardado", true, file);
                LimpiarProcesos();

            }
            else
            {
                //Debugger.Launch();
                Thread.Sleep(1000);
                Session = RootSession();
                Session = ReloadSession1(Session, "Shell_TrayWnd");
                Thread.Sleep(1000);
                var venExcel = Session.FindElementByName("Excel - 1 ventana de ejecución");
                if (venExcel != null)
                {
                    venExcel.Click();
                    Thread.Sleep(4000);
                    int count = 0;
                    while (count < 240)
                    {
                        try
                        {
                            Session = RootSession();
                            Session = ReloadSession1(Session, "XLMAIN");
                            Thread.Sleep(500);
                            var saveExcel1 = Session.FindElementsByName("Maximizar");
                            if (saveExcel1.Count > 0)
                            {
                                break;
                            }
                        }
                        catch
                        {
                            Thread.Sleep(1000);
                            count++;
                        }
                    }
                    var saveExcel = Session.FindElementsByName("Maximizar");
                    var cantidad = saveExcel.Count;
                    if (cantidad == 2)
                    {
                        saveExcel[1].Click();
                    }
                    else
                    {
                        saveExcel[0].Click();
                    }
                    //Cambio Jose
                    Session.FindElementByName("Pestaña Archivo").Click();
                    //Fin Cambio Jose

                    //Session.FindElementByName("Guardar").Click();
                    Session.FindElementByName("Guardar").Click();
                    Session.FindElementByName("Examinar").Click();
                    Thread.Sleep(1000);
                    Session.Keyboard.SendKeys(ruta);
                    Session.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo("Reporte Excel Guardado", true, file);
                    //LimpiarProcesos();
                    Thread.Sleep(2000);
                }
            }

            Thread.Sleep(2000);

        }

    }
}