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
    class CrudKbisolqu : FuncionesVitales
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
        public static void CRUDKBISolqu(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            List<string> data = new List<string>() { crudVars[0], crudVars[1], crudVars[2], crudVars[3] };
            if (tipo == 1)
            {
                ElementList[2].SendKeys(crudVars[6]);
                PruebaCRUD.LupaDinamica(desktopSession, data);
                ElementList[3].SendKeys(crudVars[4]);
            }
            else
            {
                ElementList[3].Clear();
                ElementList[3].SendKeys(crudVars[5]);
            }
        }

        public static void CRUDDetalle1KBISolqu(WindowsDriver<WindowsElement> desktopSession, List<string> crudVarsdet1, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBMemo");
            ElementList[0].Clear();
            ElementList[0].SendKeys(crudVarsdet1[0]);
        }

        public static void CRUDDetalle2KBISolqu(WindowsDriver<WindowsElement> desktopSession, List<string> crudVarsdet2, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBMemo");
            ElementList[0].Clear();
            ElementList[0].SendKeys(crudVarsdet2[0]);
        }

        public static void CRUDDetalle3KBISolqu(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVarsdet2, string file)
        {

            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            WindowsDriver<WindowsElement> rootSession = null;
            if (tipo == 1)
            {
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 50, 30);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 50, 30);
                desktopSession.Mouse.Click(null);
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                desktopSession.Keyboard.SendKeys(crudVarsdet2[0]);
                Thread.Sleep(500);

                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Control + OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Control + OpenQA.Selenium.Keys.Enter);
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TFrmDesComun");
                Thread.Sleep(500);
                var ed = rootSession.FindElementsByClassName("TDBMemo");
                var boton = rootSession.FindElementsByName("Aceptar");
                Thread.Sleep(500);
                ed[0].SendKeys(crudVarsdet2[1]);
                Thread.Sleep(500);
                rootSession.Mouse.MouseMove(boton[0].Coordinates, 35, 15);
                rootSession.Mouse.Click(null);

                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 50, 30);
                desktopSession.Mouse.Click(null);
                //rootSession = RootSession();
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Control + OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Control + OpenQA.Selenium.Keys.Enter);
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TFrmDesComun2");
                Thread.Sleep(500);
                ed = rootSession.FindElementsByClassName("TDBMemo");
                boton = rootSession.FindElementsByName("Aceptar");
                ed[0].SendKeys(crudVarsdet2[2]);
                Thread.Sleep(500);
                rootSession.Mouse.MouseMove(boton[0].Coordinates, 35, 15);
                rootSession.Mouse.Click(null);

                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 50, 30);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudVarsdet2[3]);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudVarsdet2[4]);
            }
            else
            {
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 50, 30);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                desktopSession.Keyboard.SendKeys(crudVarsdet2[5]);
            }
        }

        public static List<string> PreliKBISolqu(WindowsDriver<WindowsElement> desktopSession, string file, string pdf1, string TipoQbe, string QbeFilter, string BanderaPreli)
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
            rootSession = ReloadSession(rootSession, "TFrmTipRepo");
            var bot = rootSession.FindElementsByClassName("TGroupButton");
            bot[0].Click();
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Opción Preliminar", true, file);
            Thread.Sleep(500);
            rootSession = RootSession();
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
            if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
            for (int i = 1; i <= 2; i++)
            {
                //FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
                Thread.Sleep(2000);
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TFrmTipRepo");
                bot = rootSession.FindElementsByClassName("TGroupButton");
                Thread.Sleep(1000);
                bot[i].Click();
                Thread.Sleep(500);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                newFunctions_4.ScreenshotNuevo("Opción Preliminar", true, file);
                Thread.Sleep(1000);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);
                errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                /*if (i==1)
                {
                    string navigator = AbrirPrograma.SelectNavigator();
                    if (navigator == "IE")
                    {
                        rootSession = PruebaCRUD.RootSession();
                        rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
                        var allFrame = rootSession.FindElementsByClassName("IEFrame");
                        new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) + 15, (allFrame[0].Size.Height / 2) + 120).DoubleClick().Click().Perform();
                        Thread.Sleep(1000);
                    }
                    else
                    {
                        rootSession = PruebaCRUD.RootSession();
                        //rootSession = PruebaCRUD.ReloadSession(rootSession, "#32769");
                        //var allFrame = rootSession.FindElementsByClassName("#32769");
                        //new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) + 15, (allFrame[0].Size.Height / 2) + 120).DoubleClick().Click().Perform();
                        Thread.Sleep(1000);
                    }

                }
                else
                {
                    errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                    if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                }*/
            }
            return errorMessages;
        }


        static public void reportePreliminar(WindowsDriver<WindowsElement> desktopSession, int bandera, string file)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TGridPanel");
            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 130, 17);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);

            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession = ReloadSession(RootSession, "TFrmTipRepo");

            var btnTDVI2 = RootSession.FindElementsByClassName("TFrmTipRepo");

            switch(bandera)
            {
                case 0:
                    newFunctions_4.ScreenshotNuevo("Primera opcion", true, file);
                    RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    break;
                case 1:
                    RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 34, 89);
                    RootSession.Mouse.Click(null);
                    newFunctions_4.ScreenshotNuevo("segunda opcion", true, file);
                    RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

                    break;
                case 2:
                    RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 150, 52);
                    RootSession.Mouse.Click(null);
                    newFunctions_4.ScreenshotNuevo("Tercera opcion", true, file);
                    RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    break;
                case 3:
                    RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 150, 89);
                    RootSession.Mouse.Click(null);
                    newFunctions_4.ScreenshotNuevo("Cuarta opcion", true, file);
                    RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    break;
       
            }
         


        }

        static public void Click(WindowsDriver<WindowsElement> desktopSession)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TKactusDBgrid");
            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 150, 34);
            desktopSession.Mouse.DoubleClick(null);
            Thread.Sleep(2000);

           

        }

    }
}

