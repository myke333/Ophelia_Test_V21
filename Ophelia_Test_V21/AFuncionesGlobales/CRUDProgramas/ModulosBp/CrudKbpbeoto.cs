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
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using OpenQA.Selenium;
using System.Drawing;
using System.Windows.Forms;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBp
{
    class CrudKbpbeoto : FuncionesVitales
    {
        public static void CRUDKBpBeoto(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            //for (int i = 0; i < ElementList.Count; i++)
            //{
            //    ElementList[i].Click();
            //    Thread.Sleep(500);
            //}
            WindowsDriver<WindowsElement> rootSession = null;
            List<string> data = new List<string>() { crudVars[0], crudVars[1], "" };
            if (tipo == 1)
            {
                LupaDinamica(desktopSession, data);
                //FUNCION CERRAR VENTANA EMERGENTE EN EDGE
                string navigator = AbrirPrograma.SelectNavigator();
                if (navigator == "IE")
                {
                    rootSession = PruebaCRUD.RootSession();
                    rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
                    var allFrame = rootSession.FindElementsByClassName("IEFrame");
                    new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2), (allFrame[0].Size.Height / 2) + 120).DoubleClick().Click().Perform();
                    desktopSession.Mouse.Click(null);
                }

                Thread.Sleep(1000);

            }
            else
            {
                ElementList[6].Clear();
                ElementList[6].SendKeys(crudVars[2]);
            }
        }
        
        public static void CRUDDetalle2Kbpbeoto(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVarsdet2, string file)
        {
            WindowsDriver<WindowsElement> rootSession = null;
            string navigator = AbrirPrograma.SelectNavigator();
            if (navigator == "IE")
            {
                rootSession = PruebaCRUD.RootSession();
                rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
                var allFrame = rootSession.FindElementsByClassName("IEFrame");
                new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) - 20, (allFrame[0].Size.Height / 2) + 40).DoubleClick().Click().Perform();
                desktopSession.Mouse.Click(null);
            }
            else
            {
                rootSession = PruebaCRUD.RootSession();
                var allFrame = rootSession.FindElementsByClassName("#32769");
                Thread.Sleep(2000);
                new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) - 20, (allFrame[0].Size.Height / 2) + 40).DoubleClick().Click().Perform();
                rootSession.Mouse.Click(null);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                /*
                rootSession = PruebaCRUD.RootSession();
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);*/
            }
                

            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 25, 25);
            desktopSession.Mouse.Click(null);
            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
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

        public static void LupaDinamica(WindowsDriver<WindowsElement> desktopSession, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            List<Point> coordenadas = PruebaCRUD.CoordinatesFinder(desktopSession, 82, 150, 75);
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = PruebaCRUD.RootSession();
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            int contLupa = 1;
            bool IsDisplayedQbe = false;
            int indexVal = 0;
            foreach (Point coord in coordenadas)
            {
                if (indexVal == 2)
                {
                    break;
                }
                Thread.Sleep(2000);
                //Actions noteClicks = new Actions(desktopSession);
                //noteClicks.MoveToElement(parentElement).MoveByOffset(coord.X + 10, coord.Y).Click().Perform();

                desktopSession.Mouse.MouseMove(parentElement.Coordinates, coord.X + 10, coord.Y);
                desktopSession.Mouse.Click(null);
                if (contLupa == 1)
                {
                    try
                    {
                        rootSession = PruebaCRUD.RootSession();
                        rootSession = PruebaCRUD.ReloadSession(rootSession, "TKQBE");
                    }
                    catch
                    {

                    }
                    if (rootSession != null)
                    {
                        IsDisplayedQbe = true;
                    }
                    else
                    {
                        errorMessages.Add($"No puede encontrar la ventana de QBE");
                    }
                    if (IsDisplayedQbe)
                    {
                        var elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                        if (!string.IsNullOrEmpty(data[indexVal]))
                        {
                            elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/Pane[contains(@ClassName, 'TStringGrid')]/*"));
                            elements[0].SendKeys(data[indexVal]);
                        }
                        elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                        new Actions(rootSession).MoveToElement(elements[8], 0, 0).MoveByOffset(20, 10).DoubleClick().Perform();
                        Thread.Sleep(500);
                        IsDisplayedQbe = false;

                    }
                    Thread.Sleep(1000);
                    rootSession = PruebaCRUD.RootSession();
                    rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");
                    //var campo = rootSession.FindElementsByXPath("//Window[@ClassName=\"TCaptura\"][@Name=\" LISTA DE EMPLEADOS\"]/Edit[@ClassName=\"TEdit\"]");
                    var campo = rootSession.FindElementsByClassName("TEdit");
                    ////Debugger.Launch();
                    rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                    rootSession.Mouse.Click(null);
                    Thread.Sleep(1500);
                    rootSession.Keyboard.SendKeys(data[indexVal]);
                    //Aqui se puede enviar un valor

                    var btn = rootSession.FindElementsByClassName("TBitBtn");
                    //Debugger.Launch();
                    foreach (var boton in btn)
                    {
                        if (boton.Text == "Aceptar")
                        {
                            rootSession.Mouse.MouseMove(boton.Coordinates, 10, 10);
                            rootSession.Mouse.Click(null);
                        }
                    }
                    indexVal++;
                }
            }

        }

        public static List<string> PreliminarBpBeoto(WindowsDriver<WindowsElement> desktopSession, string BanderaPreli, string file, string pdf1, string nomPrograma)
        {
            List<string> ErrorValidacion = new List<string>();
            List<string> errors = new List<string>();
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = PruebaCRUD.RootSession();
            rootSession = PruebaCRUD.ReloadSession(rootSession, "TFrmMenuRepor");
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2500);
            errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1, nomPrograma);
            if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
            Thread.Sleep(2500);

            pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
            errors = FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2500);
            errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1, nomPrograma);
            if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
            return errorMessages;
        }
    }
}
