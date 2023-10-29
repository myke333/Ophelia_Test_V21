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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSo
{
    class CrudKsoviepi : FuncionesVitales
    {
        public static void CRUDKSoViepi(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            if (tipo == 1)
            {
                List<string> data = new List<string>() { crudVars[0] };
                PruebaCRUD.LupaDinamica(desktopSession, data);
            }
            else
            {
                List<string> data = new List<string>() { crudVars[1] };
                PruebaCRUD.LupaDinamica(desktopSession, data);
            }
        }

        public static void CRUDDetalle1KSoViepi(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVarsdet1, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
           
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 199, 183);
            desktopSession.Mouse.DoubleClick(null);
            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 54, 49);
            desktopSession.Mouse.Click(null);
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 175, 32);
            desktopSession.Mouse.DoubleClick(null); 
            var ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
            if (tipo == 1)
            {
                desktopSession.Mouse.MouseMove(ElementList2[0].Coordinates, 179, 12);
                desktopSession.Mouse.Click(null);
                desktopSession.Mouse.Click(null);
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                WindowsDriver<WindowsElement> RootSession2 = null;
                RootSession2 = PruebaCRUD.RootSession();
                RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
                ElementList2[0].SendKeys(crudVarsdet1[0]);
                for (int i = 1; i <= 2; i++)
                {
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
                    ElementList2[0].SendKeys(crudVarsdet1[i]);
                }
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
                ElementList2[0].SendKeys(crudVarsdet1[3]);
            }
            else
            {
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 420, 30);
                desktopSession.Mouse.Click(null);
                ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
                if (ElementList2.Count == 0)
                {
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
                }
                ElementList2[0].Clear();
                ElementList2[0].SendKeys(crudVarsdet1[4]);
            }            
        }

        public static List<string> PreliKSoViepi(WindowsDriver<WindowsElement> desktopSession, string file, string pdf1, string TipoQbe, string QbeFilter, string BanderaPreli,string FechaPreli)
        {

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
            List<string> errorMessages = new List<string>();
            try
            {
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TFrmMenu");
                var bot = rootSession.FindElementsByClassName("TGroupButton");
                bot[4].Click();
                newFunctions_4.ScreenshotNuevo("Tipo reporte", true, file);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);
                newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                Thread.Sleep(1000);
                FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
                bot[3].Click();
                var TEdit = rootSession.FindElementsByClassName("TEdit");
                TEdit[1].SendKeys(FechaPreli);
                TEdit[0].SendKeys(FechaPreli);
                newFunctions_4.ScreenshotNuevo("Tipo reporte", true, file);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);
                newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                Thread.Sleep(1000);
                for (int i = 2; i >= 0; i--)
                {
                    FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                    pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                    FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
                    bot[i].Click();
                    newFunctions_4.ScreenshotNuevo("Tipo reporte", true, file);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);
                    newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                    Thread.Sleep(1000);
                }
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Error de Appium: {ex.ToString()}");
            }
            return errorMessages;
            
            

        }

    }
}
