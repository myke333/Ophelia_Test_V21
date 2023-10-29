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
    class CrudKsomaobs : FuncionesVitales
    {
        public static void CRUDKSoMaobs(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList2 = desktopSession.FindElementsByClassName("TDBMemo");            
            List<string> data = new List<string>() { crudVars[0], crudVars[1], crudVars[2], crudVars[3], crudVars[4] };
            if (tipo == 1)
            {
                PruebaCRUD.LupaDinamica(desktopSession, data);
                ElementList[20].SendKeys(crudVars[5]);
                ElementList2[3].SendKeys(crudVars[8]);
                ElementList2[2].SendKeys(crudVars[6]);
                ElementList2[1].SendKeys(crudVars[9]);
                ElementList2[0].SendKeys(crudVars[7]);
            }
            else
            {
                ElementList2[1].Clear();
                ElementList2[1].SendKeys(crudVars[10]);
            }
        }
        public static void CRUDDetalle1KSoMaobs(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVarsdet1, string file)
        {
            WindowsDriver<WindowsElement> rootSession = null;
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 60, 30);
            desktopSession.Mouse.Click(null);
            var ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
            if (tipo == 1)
            {
                for (int i = 0; i <= 2; i++)
                {
                    ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
                    if (ElementList2.Count == 0)
                    {
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
                    }
                    ElementList2[0].SendKeys(crudVarsdet1[i]);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                }            
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

        public static List<string> PreliKSoMaobs(WindowsDriver<WindowsElement> desktopSession, string file, string pdf1, string TipoQbe, string QbeFilter, string BanderaPreli)
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
            string pdf = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
            try
            {
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TFrmOpcRepo");
                var bot = rootSession.FindElementsByClassName("TGroupButton");
                bot[1].Click();
                newFunctions_4.ScreenshotNuevo("Tipo reporte", true, file);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);
                newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                Thread.Sleep(1000);
                FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
                bot[0].Click();
                newFunctions_4.ScreenshotNuevo("Tipo reporte", true, file);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);
                newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                Thread.Sleep(1000);
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Error de Appium: {ex.ToString()}");
            }
            return errorMessages;           
           
        }
    }
}
