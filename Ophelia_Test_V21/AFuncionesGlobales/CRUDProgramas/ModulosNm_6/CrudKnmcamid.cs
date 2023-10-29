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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_6
{
    class CrudKnmcamid : FuncionesVitales
    {
        public static void CRUDKNmCamid(WindowsDriver<WindowsElement> desktopSession, List<string> crudVars, string file)
        {
            List<string> data = new List<string>() { crudVars[0]  };
            var Element = desktopSession.FindElementsByClassName("TCheckBox");
            var Element2 = desktopSession.FindElementsByName("Tarjeta Identidad");
            var Aceptar = desktopSession.FindElementsByName("Aceptar");
            for (int i = 0; i < 2; i++)
            {
                Thread.Sleep(500);
                LupaDinamica(desktopSession, data);
                Thread.Sleep(500);
                WindowsDriver<WindowsElement> rootSession = null;
                rootSession = PruebaCRUD.RootSession();
                rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
                Element = rootSession.FindElementsByClassName("TCheckBox");
                Element2 = rootSession.FindElementsByName("Tarjeta Identidad");
                Aceptar = rootSession.FindElementsByName("Aceptar");
                Element[0].Click();
                Thread.Sleep(500);
                Element2[0].Click();
                Thread.Sleep(500);
                Aceptar[0].Click();
                Thread.Sleep(500);
                rootSession = PruebaCRUD.RootSession();
                rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
                Thread.Sleep(500);
                newFunctions_4.ScreenshotNuevo("Proceso Finalizado", true, file);
                var allFrame = rootSession.FindElementsByClassName("IEFrame");
                Thread.Sleep(500);
                new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) + 90, (allFrame[0].Size.Height / 2) + 80).DoubleClick().Perform();
                Thread.Sleep(500);
                Element2 = desktopSession.FindElementsByName("Cédula");
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
                Thread.Sleep(2000);
                //Actions noteClicks = new Actions(desktopSession);
                //noteClicks.MoveToElement(parentElement).MoveByOffset(coord.X + 10, coord.Y).Click().Perform();
                desktopSession.Mouse.MouseMove(parentElement.Coordinates, coord.X + 10, coord.Y);
                desktopSession.Mouse.Click(null);

                rootSession = PruebaCRUD.RootSession();
                rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");

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

                rootSession = PruebaCRUD.RootSession();
                rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
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

