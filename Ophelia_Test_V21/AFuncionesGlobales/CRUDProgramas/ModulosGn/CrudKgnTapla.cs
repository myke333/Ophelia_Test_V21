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
using System.Windows.Forms;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosGn
{
    class CrudKgnTapla : FuncionesVitales
    {
        static public WindowsDriver<WindowsElement> RootSession()
        {
            var appCapabilities = new AppiumOptions();
            appCapabilities.AddAdditionalCapability("app", "Root");
            var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
            return ApplicationSession;
        }

        
        static public WindowsDriver<WindowsElement> ReloadSession1(WindowsDriver<WindowsElement> session, string parametro)
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

        static public void CrudTapla(WindowsDriver<WindowsElement> desktopSession, int tipo)
        {

            var Elementlist = desktopSession.FindElementsByClassName("TEdit");
            var Elementlist1 = desktopSession.FindElementsByClassName("TScrollBox");
            var Elementlist2 = desktopSession.FindElementsByClassName("TPanel");
            if (tipo == 0)
            {


                desktopSession.Mouse.MouseMove(Elementlist[1].Coordinates, 360, 120);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);


                //desktopSession.Mouse.MouseMove(Elementlist2[1].Coordinates, 360, 32);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1500);

                //WindowsDriver<WindowsElement> RootSession = null;
                //RootSession = PruebaCRUD.RootSession();
                //RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                //Thread.Sleep(2000);



            }
            else
            {

                desktopSession.Mouse.MouseMove(Elementlist[1].Coordinates, 360, 197);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);

               


            }
        }

        static public void BotonAceptar(WindowsDriver<WindowsElement> desktopSession, int tipo)
        {

            //var Elementlist = desktopSession.FindElementsByClassName("TEdit");
            var Elementlist1 = desktopSession.FindElementsByClassName("TScrollBox");
            //var Elementlist2 = desktopSession.FindElementsByClassName("TPanel");
            if (tipo == 0)
            {
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(Elementlist1[0].Coordinates, 373, 545);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);

                





            }
            else
            {
                Thread.Sleep(2000);
                WindowsDriver<WindowsElement> RootSession = null;
                RootSession = PruebaCRUD.RootSession();
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);



            }
        }

        //static public void ExpExel(WindowsDriver<WindowsElement> Session, string ruta)
        //{




        //        //Abre ventana excel
        //        Session = RootSession();
        //        Session = ReloadSession1(Session, "XLMAIN");
        //        if (Session != null)
        //        {
        //            int count = 0;
        //            while (count < 240)
        //            {
        //                try
        //                {
        //                    //Session = RootSession();
        //                    //Session = ReloadSession1(Session, "XLMAIN");
        //                    Thread.Sleep(500);
        //                    var saveExcel1 = Session.FindElementsByName("Maximizar");
        //                    if (saveExcel1.Count > 0)
        //                    {
        //                        break;
        //                    }
        //                }
        //                catch
        //                {
        //                    Thread.Sleep(1000);
        //                    count++;
        //                }
        //            }
        //            var saveExcel = Session.FindElementsByName("Maximizar");
        //            var cantidad = saveExcel.Count;
        //            if (cantidad == 2)
        //            {
        //                saveExcel[1].Click();
        //            }
        //            else
        //            {
        //                saveExcel[0].Click();
        //            }
        //            //Cambio Jose
        //            Session.FindElementByName("Pestaña Archivo").Click();
        //            //Fin Cambio Jose



        //            //Session.FindElementByName("Guardar").Click();
        //            Session.FindElementByName("Guardar").Click();
        //            Session.FindElementByName("Examinar").Click();
        //            Thread.Sleep(1000);
        //            Session.Keyboard.SendKeys(ruta);
        //            Session.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
        //            Thread.Sleep(1000);
        //            //newFunctions_4.ScreenshotNuevo("Reporte Excel Guardado", true, file);
        //            LimpiarProcesos();



        //        }
        //        else
        //        {
        //            //Debugger.Launch();
        //            Thread.Sleep(1000);
        //            Session = RootSession();
        //            Session = ReloadSession1(Session, "Shell_TrayWnd");
        //            Thread.Sleep(1000);
        //            var venExcel = Session.FindElementByName("Excel - 1 ventana de ejecución");
        //            if (venExcel != null)
        //            {
        //                venExcel.Click();
        //                Thread.Sleep(4000);
        //                int count = 0;
        //                while (count < 240)
        //                {
        //                    try
        //                    {
        //                        Session = RootSession();
        //                        Session = ReloadSession1(Session, "XLMAIN");
        //                        Thread.Sleep(500);
        //                        var saveExcel1 = Session.FindElementsByName("Maximizar");
        //                        if (saveExcel1.Count > 0)
        //                        {
        //                            break;
        //                        }
        //                    }
        //                    catch
        //                    {
        //                        Thread.Sleep(1000);
        //                        count++;
        //                    }
        //                }
        //                var saveExcel = Session.FindElementsByName("Maximizar");
        //                var cantidad = saveExcel.Count;
        //                if (cantidad == 2)
        //                {
        //                    saveExcel[1].Click();
        //                }
        //                else
        //                {
        //                    saveExcel[0].Click();
        //                }
        //                //Cambio Jose
        //                Session.FindElementByName("Pestaña Archivo").Click();
        //                //Fin Cambio Jose



        //                //Session.FindElementByName("Guardar").Click();
        //                Session.FindElementByName("Guardar").Click();
        //                Session.FindElementByName("Examinar").Click();
        //                Thread.Sleep(1000);
        //                Session.Keyboard.SendKeys(ruta);
        //                Session.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
        //                Thread.Sleep(1000);
        //                //newFunctions_4.ScreenshotNuevo("Reporte Excel Guardado", true, file);
        //                //LimpiarProcesos();
        //                Thread.Sleep(2000);
        //            }
        //        }



        //        Thread.Sleep(2000);

        //    }
        //static public void Boton(WindowsDriver<WindowsElement> Session)
        //{
        //    Session.FindElementByName("exel").Click();
        //    Thread.Sleep(2000);



        //    //Abre ventana export excel



        //    WindowsDriver<WindowsElement> RootSession2 = null;
        //    RootSession2 = PruebaCRUD.RootSession();



        //    RootSession2 = ReloadSession(RootSession2, "TFrmExcelex");
        //    var btnTDVI2 = RootSession2.FindElementsByClassName("TFrmExcelex");
        //    // Se escogen todos los >> atributos a exportar
        //    RootSession2.Mouse.MouseMove(btnTDVI2[0].Coordinates, 221, 210);
        //    RootSession2.Mouse.Click(null);
        //    Thread.Sleep(1000);



        //    RootSession2.FindElementByName("Aceptar").Click();
        //    Thread.Sleep(6000);



        //    RootSession2.Mouse.Click(null);
        //    Thread.Sleep(2000);








        /*

        var cont = 0;

        foreach (var item in Elementlist)
        {
            if (cont == 3)
            {
                item.SendKeys("2");
            }
            else
            {
                item.SendKeys(Convert.ToString(cont));
            }

            cont = cont + 1;
            Thread.Sleep(1000);
        }
        */


    }
}

