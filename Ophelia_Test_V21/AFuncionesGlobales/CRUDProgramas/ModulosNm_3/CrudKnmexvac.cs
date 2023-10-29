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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_3
{
    class CrudKnmexvac : FuncionesVitales
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



        public static void CRUDKNmExvac(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {

            var ElementList = desktopSession.FindElementsByClassName("TDBEdit"); 
            //for (int i = 0; i < ElementList.Count; i++)
            //{
            //    ElementList[i].Click();
            //    Thread.Sleep(500);
            //}
            if (tipo == 1)
            {
                List<string> data = new List<string>() { crudVars[0], crudVars[1] };
                PruebaCRUD.LupaDinamica(desktopSession, data);

                ElementList[8].SendKeys(crudVars[2]);
                ElementList[7].SendKeys(crudVars[3]);

                ElementList[9].SendKeys(crudVars[9]);
                ElementList[10].SendKeys(crudVars[8]);
                ElementList[11].SendKeys(crudVars[7]);

                ElementList[16].SendKeys(crudVars[4]);
                ElementList[15].SendKeys(crudVars[5]);
                ElementList[14].SendKeys(crudVars[6]);
            }
            else
            {
                ElementList[7].Clear();
                ElementList[7].SendKeys(crudVars[10]);
            }
        }

        static public void Exceldistinto(WindowsDriver<WindowsElement> desktopSession, int bandera, string file)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TGridPanel");
            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 574, 14);
            desktopSession.Mouse.Click(null);

            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession = ReloadSession(RootSession, "TForm");


            var btnTDVI2 = RootSession.FindElementsByClassName("TForm");


            //TPanel
            if (bandera == 0)
            {

                newFunctions_4.ScreenshotNuevo("Primera opcion", true, file);
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);

            }
            else
            {
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                newFunctions_4.ScreenshotNuevo("Segunda opcion", true, file);
                Thread.Sleep(2000);
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);

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
