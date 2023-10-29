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
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm
{
    class CrudKnmcuent
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

        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data,string file)
        {
            List<string> errorMessages = new List<string>();

            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList3 = desktopSession.FindElementsByClassName("TGroupButton");
            if (tipo == 1)
            {
                newFunctions_4.ScreenshotNuevo("Agregar Registro", true, file);
                foreach (var elem in ElementList3)
                {
                    if (elem.Text == "(CCF) Caja de Compensacion Familiar")
                    {
                        elem.Click();
                    }
                }
                LupaDinamica(desktopSession, data);
                Thread.Sleep(1000);
                ElementList[8].SendKeys(data[3]);
            }
            else
            {
                newFunctions_4.ScreenshotNuevo("Editar Registro", true, file);
                ElementList[8].Clear();
                ElementList[8].SendKeys(data[4]);
            }
        }
        public static List<string> Preliminar(WindowsDriver<WindowsElement> desktopSession, string file, string pdf1, string BanderaPreli)
        {
            //FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            try
            {            
            rootSession = RootSession();
            rootSession = ReloadSession(rootSession, "TFrmTipoRep");

            var ElementList = rootSession.FindElementsByClassName("TGroupButton");
            var ElementList2 = rootSession.FindElementsByClassName("TGroupButton");

            var boton = rootSession.FindElementsByClassName("TBitBtn");

            List<string> rep = new List<string> { "Activo", "Inacivo", "Pensionado", "Jubilado", "Retirado", "Beneficiario de Pensión", "Todos" };

            for (int i = 0; i < 7; i++)
            {
                FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TFrmTipoRep");
                ElementList = rootSession.FindElementsByClassName("TGroupButton");


                foreach (var elem in ElementList)
                {
                    if (elem.Text == "Resumido")
                    {
                        elem.Click();
                    }
                }
                ElementList2 = rootSession.FindElementsByClassName("TGroupButton");
                foreach (var L in ElementList2)
                {
                    if (L.Text == rep[i])
                    {
                        L.Click();
                    }
                }
                boton = rootSession.FindElementsByClassName("TBitBtn");
                foreach (var btn in boton)
                {
                    if (btn.Text == "Aceptar")
                    {
                        btn.Click();
                        newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                    }
                }

            }
            for (int i = 0; i < 7; i++)
            {
                FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TFrmTipoRep");
                ElementList = rootSession.FindElementsByClassName("TGroupButton");


                foreach (var elem in ElementList)
                {
                    if (elem.Text == "Detallado")
                    {
                        elem.Click();
                    }
                }
                ElementList2 = rootSession.FindElementsByClassName("TGroupButton");
                foreach (var L in ElementList2)
                {
                    if (L.Text == rep[i])
                    {
                        L.Click();
                    }
                }
                boton = rootSession.FindElementsByClassName("TBitBtn");
                foreach (var btn in boton)
                {
                    if (btn.Text == "Aceptar")
                    {
                        btn.Click();
                        newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                    }
                }

            }
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Error de Appium: {ex.ToString()}");
            }
            return errorMessages;

        }
        public static void LupaDinamica(WindowsDriver<WindowsElement> desktopSession, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            List<Point> coordenadas = PruebaCRUD.CoordinatesFinder(desktopSession, 82, 150, 75);
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            int contLupa = 1;
            bool IsDisplayedQbe = false;
            int indexVal = 0;
            int contqbe = 0;
            foreach (Point coord in coordenadas)
            {
                Thread.Sleep(2000);
                //Actions noteClicks = new Actions(desktopSession);
                //noteClicks.MoveToElement(parentElement).MoveByOffset(coord.X + 10, coord.Y).Click().Perform();

                desktopSession.Mouse.MouseMove(parentElement.Coordinates, coord.X + 10, coord.Y);
                desktopSession.Mouse.Click(null);
                if (contLupa == 1)
                {
                    if (contqbe == 0)
                    {

                        contqbe++;
                        try
                        {
                            rootSession = RootSession();
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
                    }
                    rootSession = RootSession();
                    rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");
                    //var campo = rootSession.FindElementsByXPath("//Window[@ClassName=\"TCaptura\"][@Name=\" LISTA DE EMPLEADOS\"]/Edit[@ClassName=\"TEdit\"]");
                    var campo = rootSession.FindElementsByClassName("TEdit");
                    ////Debugger.Launch();
                    rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                    rootSession.Mouse.Click(null);
                    rootSession.Keyboard.SendKeys(data[indexVal]);
                    //Aqui se puede enviar un valor

                    var btn = rootSession.FindElementsByClassName("TBitBtn");
                    ////Debugger.Launch();
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
    }
}
