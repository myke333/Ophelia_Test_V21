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
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosPn
{
    class CrudKpncppen : FuncionesVitales
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
        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, string tipo, List<String> data, string file)
        {
            var Elementlist = desktopSession.FindElementsByClassName("TDBEdit");
         
            var Elementlist1 = desktopSession.FindElementsByName(data[17]);
            var Elementlist2 = desktopSession.FindElementsByName(data[18]);
            switch (tipo)
            {
                case ("0"):
                    Elementlist[31].SendKeys(data[0]);
                  
                    Elementlist[32].SendKeys(data[1]);
                
                    Elementlist[28].SendKeys(data[2]);
           
                    Elementlist[27].SendKeys(data[3]);
       
                    Elementlist[24].SendKeys(data[4]);
             
                    Elementlist[23].SendKeys(data[5]);
       
                    Elementlist[22].SendKeys(data[6]);
             
                    Elementlist[20].SendKeys(data[7]);
                   
                    Elementlist[19].SendKeys(data[8]);
           
                    Elementlist[18].SendKeys(data[9]);
             
                    Elementlist[17].SendKeys(data[10]);
                
                    Elementlist[15].SendKeys(data[11]);
                 
                    Elementlist[14].SendKeys(data[12]);
                  
                    Elementlist[13].SendKeys(data[13]);
                 
                    Elementlist[12].SendKeys(data[14]);
                  
                    Elementlist[11].SendKeys(data[15]);
                
                    Elementlist[8].SendKeys(data[16]);
                   
                    Elementlist1[0].Click();
                    break;
                case ("1"):
                    Elementlist2[0].Click();
                    break;
            }
        }
        public static void CRUD_DETALLE(WindowsDriver<WindowsElement> desktopSession, string tipo, List<String> data, string file)
        {
            //TDBGridInplaceEdit
            var Elementlist = desktopSession.FindElementsByClassName("TDBGrid");
            switch (tipo)
            {
                case ("0"):
                    Elementlist[0].SendKeys(OpenQA.Selenium.Keys.Tab);
                    Elementlist[0].SendKeys(OpenQA.Selenium.Keys.Tab);
                    desktopSession.Keyboard.SendKeys(data[0]);
                    break;
                case ("1"):
                    desktopSession.Keyboard.SendKeys(data[1]);
                    break;
            }
        }

        public static List<string> preliminarDiferente(WindowsDriver<WindowsElement> desktopSession, string BanderaPreli, string file, string nomPrograma)
        {
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();
            for (int i = 0; i < 3; i++)
            {
                string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                errors = FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

                Thread.Sleep(2000);

                WindowsDriver<WindowsElement> rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TFrmListReports");
                
                switch (i)
                {
                    case 0:
                        //rootSession.FindElementByName("Reporte General Cuotas Partes Pensionales").Click();
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(1000);
                        newFunctions_4.ScreenshotNuevo("Selección Opción Preliminar", true, file);
                        Thread.Sleep(1000);
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        Thread.Sleep(10000);
                        errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1, nomPrograma);
                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                        break;
                    case 1:
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(1000);
                        newFunctions_4.ScreenshotNuevo("Selección Opción Preliminar", true, file);
                        Thread.Sleep(1000);
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        Thread.Sleep(10000);
                        errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1, nomPrograma);
                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                        break;
                    case 2:
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        Thread.Sleep(1000);
                        newFunctions_4.ScreenshotNuevo("Selección Opción Preliminar", true, file);
                        Thread.Sleep(1000);
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        Thread.Sleep(6000);
                        rootSession = RootSession();
                        rootSession = ReloadSession(rootSession, "TFrmFirmas");
                        Thread.Sleep(1500);
                        rootSession.Keyboard.SendKeys("Prueba");
                        Thread.Sleep(1000);
                        newFunctions_4.ScreenshotNuevo("Selección Opción Firmas Preliminar", true, file);
                        Thread.Sleep(1000);
                        rootSession.FindElementByName("Aceptar").Click();
                        Thread.Sleep(10000);
                        errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1, nomPrograma);
                        if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
                        break;

                    default:
                        break;
                }
            }
            return ErrorValidacion;
        }
    }
}
