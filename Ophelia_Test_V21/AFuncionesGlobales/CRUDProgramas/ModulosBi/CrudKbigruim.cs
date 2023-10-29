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
    class CrudKbigruim : FuncionesVitales
    {
        static public WindowsDriver<WindowsElement> RootSession()
        {

            var appCapabilities = new AppiumOptions();
            appCapabilities.AddAdditionalCapability("app", "Root");
            var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
            return ApplicationSession;
        }

        static public WindowsDriver<WindowsElement> ReloadSession(WindowsDriver<WindowsElement> session, string className)
        {
            try
            {
                var ApplicationWindow = session.FindElementByClassName(className);
                var ApplicationSessionHandle = ApplicationWindow.GetAttribute("NativeWindowHandle");
                ApplicationSessionHandle = (int.Parse(ApplicationSessionHandle)).ToString("x"); //Convert to hex


                var appCapabilities = new AppiumOptions();
                //appCapabilities.AddAdditionalCapability("app", "Root");
                appCapabilities.AddAdditionalCapability("appTopLevelWindow", ApplicationSessionHandle);
                var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);

                string ses = ApplicationSession.PageSource;
                ////Debugger.Launch();
                return ApplicationSession;
            }
            catch (Exception ex) { return null; }
        }

        public static void CRUDKBiGruim(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            
            ElementList[2].SendKeys(crudVars[0]);
            ElementList[3].SendKeys(crudVars[1]);           
        }

        public static void CRUDDetalle1KBiGruim(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVarsdet1, string file)
        {
                      
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates,135, 30);
            desktopSession.Mouse.Click(null);
            desktopSession.Mouse.Click(null);
            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Control + OpenQA.Selenium.Keys.Enter);
            WindowsDriver<WindowsElement> rootSession = null;

            rootSession = PruebaCRUD.RootSession();
            rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");
            var ElementList2 = rootSession.FindElementsByClassName("TEdit");
            
            if (tipo == 1)
            {                
                ElementList2[0].SendKeys(crudVarsdet1[0]);
                ElementList2[1].SendKeys(crudVarsdet1[0]);
            }
            else
            {
                ElementList2[0].Clear();
                ElementList2[0].SendKeys(crudVarsdet1[1]);
                ElementList2[1].Clear();
                ElementList2[1].SendKeys(crudVarsdet1[1]);
            }
            var ElementList3 = rootSession.FindElementsByName("Aceptar");
            rootSession.Mouse.MouseMove(ElementList3[0].Coordinates, 31, 14);
            rootSession.Mouse.Click(null);

        }

        public static List<string> PreliKBiGruim(WindowsDriver<WindowsElement> desktopSession, string file, string pdf1, string TipoQbe, string QbeFilter, string BanderaPreli)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            WindowsDriver<WindowsElement> RootSession()
            {
                var appCapabilities = new AppiumOptions();
                appCapabilities.AddAdditionalCapability("app", "Root");
                var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
                return ApplicationSession;
            }
            WindowsDriver<WindowsElement> ReloadSession(WindowsDriver<WindowsElement> session, string parametro)
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
            rootSession = RootSession();
            rootSession = ReloadSession(rootSession, "TTipoReporeFRM");
            var bot = rootSession.FindElementsByName("Aceptar");
            bot[0].Click();
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Opción Preliminar", true, file);
            Thread.Sleep(500);
            rootSession = RootSession();
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
            if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
            return errorMessages;
        }


        static public void CrudPrincipal(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");
            //TPanel
            if (bandera == 0)
            {

                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 23, 47);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 154, 47);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[1]);

            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 82, 111);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[2]);

            }

        }

        static public void CrudDetalle(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudVariablesDetalle1, string file)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TDBGrid");
            //TPanel
            if (bandera == 0)
            {

                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 116, 35);
                desktopSession.Mouse.DoubleClick(null);

                WindowsDriver<WindowsElement> RootSession = null;
                RootSession = PruebaCRUD.RootSession();
                RootSession = ReloadSession(RootSession, "TCaptura");

                newFunctions_4.ScreenshotNuevo("Pestaña", true, file);

                var btnTDVI2 = RootSession.FindElementsByClassName("TCaptura");



                RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 220, 69);
                RootSession.Mouse.Click(null);
                Thread.Sleep(2000);
                RootSession.Keyboard.SendKeys(crudVariablesDetalle1[0]);
                Thread.Sleep(2000);
                newFunctions_4.ScreenshotNuevo("Valor elegido", true, file);
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);

            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 116, 35);
                desktopSession.Mouse.DoubleClick(null);

                WindowsDriver<WindowsElement> RootSession = null;
                RootSession = PruebaCRUD.RootSession();
                RootSession = ReloadSession(RootSession, "TCaptura");

                newFunctions_4.ScreenshotNuevo("Pestaña", true, file);

                var btnTDVI2 = RootSession.FindElementsByClassName("TCaptura");



                RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 220, 69);
                RootSession.Mouse.Click(null);
                Thread.Sleep(2000);
                RootSession.Keyboard.SendKeys(crudVariablesDetalle1[1]);
                Thread.Sleep(2000);
                newFunctions_4.ScreenshotNuevo("Valor elegido", true, file);
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);

            }

        }


        static public void Exceldistinto(WindowsDriver<WindowsElement> desktopSession, int bandera,string file)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TGridPanel");
            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 574, 14);
            desktopSession.Mouse.Click(null);

            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession = ReloadSession(RootSession, "TFrmMenuExcel");


            var btnTDVI2 = RootSession.FindElementsByClassName("TFrmMenuExcel");


            //TPanel
            if (bandera == 0)
            {

                newFunctions_4.ScreenshotNuevo("Primera opcion", true, file);
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);

            }
            else
            {

                RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 32, 108);
                RootSession.Mouse.Click(null);
                newFunctions_4.ScreenshotNuevo("Segunda opcion", true, file);
                Thread.Sleep(2000);
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);

            }

        }


       

    }
}
