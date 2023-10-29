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
    class CrudKsopanev : FuncionesVitales
    {
        public static void CRUDKSoPanEv(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file, string motor)
        {            
            if (tipo == 1)
            {
                var ElementList = desktopSession.FindElementsByClassName("TGroupButton");
                if (motor=="SQL")
                {
                    ////PESTAÑA 1
                    List<string> data = new List<string>() { crudVars[0], crudVars[1], crudVars[2] };
                    PruebaCRUD.LupaDinamica(desktopSession, data);
                    ElementList[5].Click();
                }
                else
                {
                    ////PESTAÑA 1
                    List<string> data = new List<string>() { crudVars[0], crudVars[1], crudVars[2], crudVars[3], crudVars[4] };
                    List<int> dis = new List<int>() { 4 };
                    newFunctions_4.LupaDinamicaDiscriminatoria(desktopSession, data, dis);
                    ElementList[6].Click();
                    ElementList[7].Click();
                }
               
            }
            else
            {
                if (motor == "SQL")
                {
                    var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
                    ElementList[10].Clear();
                    ElementList[10].SendKeys(crudVars[3]);
                }
                else
                {
                    ////PESTAÑA 1
                    List<string> data = new List<string>() { crudVars[5] };
                    List<int> dis = new List<int>() { 1,2,3,4,5 };
                    newFunctions_4.LupaDinamicaDiscriminatoria(desktopSession, data, dis);
                }
               
            }
        }
        public static void CRUDDetalleKSoPanEv(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVarsdet1, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 45, 35);
            desktopSession.Mouse.Click(null);
            var ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
            if (tipo == 1)
            {
                ElementList2[0].SendKeys(crudVarsdet1[0]);
            }
            else
            {
                ElementList2[0].Clear();
                ElementList2[0].SendKeys(crudVarsdet1[1]);
            }
        }
        public static void CRUDDetalle5KSoPanEv(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVarsdet5, string file)
        {
            WindowsDriver<WindowsElement> rootSession = null;
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 25, 30);
            desktopSession.Mouse.Click(null);
            var ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
            if (tipo == 1)
            {
                ElementList2[0].SendKeys(crudVarsdet5[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudVarsdet5[1]);
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 25, 30);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Control + OpenQA.Selenium.Keys.Enter);
                rootSession = PruebaCRUD.RootSession();
                rootSession = PruebaCRUD.ReloadSession(rootSession, "TFrmCorrecion");
                var ElementList3 = rootSession.FindElementsByClassName("TDBMemo");
                Thread.Sleep(500);
                ElementList3[1].SendKeys(crudVarsdet5[2]);
                Thread.Sleep(500);
                var ElementList4 = rootSession.FindElementsByClassName("TKNavegador");
                rootSession.Mouse.MouseMove(ElementList4[0].Coordinates, 195, 17);
                rootSession.Mouse.Click(null);
                var ElementList5 = rootSession.FindElementsByName("Cerrar");
                rootSession.Mouse.MouseMove(ElementList5[0].Coordinates, 5, 5);
                rootSession.Mouse.Click(null);
            }
            else
            {
                ElementList2[0].Clear();
                ElementList2[0].SendKeys(crudVarsdet5[3]);
            }
        }
        public static void CRUDDetalle7KSoPanEv(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVarsdet7, string file)
        {
            WindowsDriver<WindowsElement> rootSession = null;
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 25, 30);
            desktopSession.Mouse.Click(null);
            var ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
            if (tipo == 1)
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Control + OpenQA.Selenium.Keys.Enter);
                rootSession = PruebaCRUD.RootSession();
                rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");
                var ElementList3 = rootSession.FindElementsByClassName("TEdit");
                ElementList3[0].SendKeys(crudVarsdet7[0]);
                var ElementList4 = rootSession.FindElementsByName("Aceptar");
                rootSession.Mouse.MouseMove(ElementList4[0].Coordinates, 10, 10);
                rootSession.Mouse.Click(null);
            }
            else
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
                ElementList2[0].Clear();
                ElementList2[0].SendKeys(crudVarsdet7[1]);
            }
        }
        public static void CRUDDetalle9KSoPanEv(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVarsdet9, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 45, 35);
            desktopSession.Mouse.Click(null);
            var ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
            if (ElementList2.Count == 0)
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
            }
            if (tipo == 1)
            {
                ElementList2[0].SendKeys(crudVarsdet9[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudVarsdet9[1]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudVarsdet9[2]);
            }
            else
            {
                ElementList2[0].Clear();
                ElementList2[0].SendKeys(crudVarsdet9[3]);
            }
        }
        public static void CRUDDetalle10KSoPanEv(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVarsdet10, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 45, 35);
            desktopSession.Mouse.Click(null);
            var ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
            if (ElementList2.Count == 0)
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
            }
            if (tipo == 1)
            {
                ElementList2[0].SendKeys(crudVarsdet10[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudVarsdet10[1]);
            }
            else
            {
                ElementList2[0].Clear();
                ElementList2[0].SendKeys(crudVarsdet10[2]);
            }
        }
        public static List<string> PreliKSoPanEv(WindowsDriver<WindowsElement> desktopSession, string file, string pdf1, string TipoQbe, string QbeFilter, string BanderaPreli,string motor)
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
                int cont = 0;
                int cont1 = 0;
                if (motor == "SQL")
                {
                    cont = 3;
                    cont1 = 1;
                }
                if (motor == "ORA")
                {
                    cont = 2;
                    cont1 = 0;
                }
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TFrmSoReporte");
                var ElementList = rootSession.FindElementsByClassName("TFrmSoReporte");
                var bot = rootSession.FindElementsByClassName("TGroupButton");
                bot[cont].Click();
                newFunctions_4.ScreenshotNuevo("Tipo reporte", true, file);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);
                newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                Thread.Sleep(1000);
                cont -= 1;
                for (int i = cont; i >= 0; i--)
                {
                    FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
                    pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                    FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
                    rootSession = RootSession();
                    rootSession = ReloadSession(rootSession, "TFrmSoReporte");
                    bot = rootSession.FindElementsByClassName("TGroupButton");
                    bot[i].Click();
                    if (i == cont1)
                    {
                        rootSession.Mouse.MouseMove(ElementList[0].Coordinates, 439, 153);
                        rootSession.Mouse.Click(null);
                        rootSession = RootSession();
                        rootSession = ReloadSession(rootSession, "TKQBE");
                        var ElementList1 = rootSession.FindElementsByClassName("TStringGrid");
                        rootSession.Mouse.MouseMove(ElementList1[0].Coordinates, 230, 60);
                        rootSession.Mouse.Click(null);
                        var ElementList2 = rootSession.FindElementsByClassName("TInplaceEdit");
                        if (ElementList2.Count == 0)
                        {
                            rootSession.Mouse.Click(null);
                        }
                        rootSession.Keyboard.SendKeys(QbeFilter);
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        Thread.Sleep(1000);
                        rootSession = RootSession();
                        rootSession = ReloadSession(rootSession, "TFrmSoReporte");
                    }
                    newFunctions_4.ScreenshotNuevo("Tipo reporte", true, file);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(2000);
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

        public static void PasarPestaña(WindowsDriver<WindowsElement> desktopSession, int Cantidad, int tipo, int numeral )
        {
            var TPageControl = desktopSession.FindElementsByClassName("TPageControl");
            desktopSession.Mouse.MouseMove(TPageControl[numeral].Coordinates, 30, 13);
            desktopSession.Mouse.Click(null);
            desktopSession.Mouse.Click(null);
            for (int i = 1; i <= Cantidad; i++)
            {
                if (tipo == 1)
                {
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
                }
                else
                {
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Left);
                }
            }
        }

    }        
}
