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
    class CrudKsomasim : FuncionesVitales
    {              
        public static void CRUDKSoMasim(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList2 = desktopSession.FindElementsByClassName("TDBMemo");
            //for (int i = 2; i <= ElementList.Count; i++)
            //{
            //    ElementList[i].Click();
            //    ElementList[i].SendKeys(""+i);
            //    Thread.Sleep(1000);
            //}
            List<string> data = new List<string>() { crudVars[0], crudVars[1], crudVars[2] };
            if (tipo == 1)
            {
                PruebaCRUD.LupaDinamica(desktopSession, data);
                ElementList[8].SendKeys(crudVars[3]);

                ElementList2[0].SendKeys(crudVars[4]);
                ElementList[15].SendKeys(crudVars[5]);
                ElementList[13].SendKeys(crudVars[6]);
                ElementList[11].SendKeys(crudVars[7]);
                ElementList[12].SendKeys(crudVars[8]);

                ElementList[21].SendKeys(crudVars[9]);
                ElementList[20].SendKeys(crudVars[10]);
                ElementList[19].SendKeys(crudVars[11]);
                ElementList[18].SendKeys(crudVars[12]);
                ElementList[17].SendKeys(crudVars[13]);
                ElementList[10].SendKeys(crudVars[14]);
            }
            else
            {
                ElementList2[0].Clear();
                ElementList2[0].SendKeys(crudVars[15]);
            }
        }
        public static void CRUDDetalle4KSoMasim(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVarsdet4, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 50, 35);
            desktopSession.Mouse.Click(null);            
            var ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
            if (tipo == 1)
            {                
                ElementList2[0].SendKeys(crudVarsdet4[0]);
            }
            else
            {
                ElementList2[0].Clear();
                ElementList2[0].SendKeys(crudVarsdet4[1]);
            }
        }
        public static void CRUDDetalle1KSoMasim(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVarsdet1, string file)
        {
            WindowsDriver<WindowsElement> rootSession = null;
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 60, 30);
            desktopSession.Mouse.Click(null);
            var ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
            if (tipo == 1)
            {                
                ElementList2[0].SendKeys(crudVarsdet1[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Control + OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(500);
                rootSession = PruebaCRUD.RootSession();
                rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");                
                var ElementList3 = rootSession.FindElementsByClassName("TEdit");
                Thread.Sleep(500);
                ElementList3[0].SendKeys(crudVarsdet1[1]);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(crudVarsdet1[2]);

            }
            else
            {
                if (ElementList2.Count == 0)
                {
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
                }
                ElementList2[0].Clear();
                ElementList2[0].SendKeys(crudVarsdet1[3]);
            }
        }
        public static void CRUDDetalle2KSoMasim(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVarsdet2, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 55, 35);
            desktopSession.Mouse.Click(null);
            var ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
            if (tipo == 1)
            {
                ElementList2[0].SendKeys(crudVarsdet2[0]);
            }
            else
            {
                ElementList2[0].Clear();
                ElementList2[0].SendKeys(crudVarsdet2[1]);
            }
        }
        public static void CRUDDetalle3KSoMasim(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVarsdet3, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, 100, 35);
            desktopSession.Mouse.Click(null);
            var ElementList2 = ElementList[1].FindElementsByClassName("TDBGridInplaceEdit");
            if (tipo == 1)
            {
                ElementList2[0].SendKeys(crudVarsdet3[0]);
            }
            else
            {
                ElementList2[0].Clear();
                ElementList2[0].SendKeys(crudVarsdet3[1]);
            }
        }
        public static void CRUDDetalle5KSoMasim(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVarsdet5, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 75, 35);
            desktopSession.Mouse.Click(null);
            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            var ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
            if (tipo == 1)
            {
                ElementList2[0].SendKeys(crudVarsdet5[0]);
            }
            else
            {
                ElementList2[0].Clear();
                ElementList2[0].SendKeys(crudVarsdet5[1]);
            }
        }
        public static List<string> PreliKSoMasim(WindowsDriver<WindowsElement> desktopSession, string file, string pdf1, string TipoQbe, string QbeFilter, string BanderaPreli)
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
                rootSession = ReloadSession(rootSession, "TFrmSoReporte");
                var bot = rootSession.FindElementsByClassName("TGroupButton");
                bot[3].Click();
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
