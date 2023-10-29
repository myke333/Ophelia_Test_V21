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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_5
{
    class CrudKnmracum : FuncionesVitales
    {
        public static void CRUDKNmRacum(WindowsDriver<WindowsElement> desktopSession, List<string> crudVars, string file)
        {
            var Element = desktopSession.FindElementsByClassName("TEdit");
            var Element1 = desktopSession.FindElementsByName("Selección");
            Thread.Sleep(500);
            Element[1].SendKeys(crudVars[0]);
            Thread.Sleep(500);
            Element1[0].Click();
        }
        public static void PreliKNmRacum(WindowsDriver<WindowsElement> desktopSession, string file)
        {
            var Element = desktopSession.FindElementsByClassName("TPanel");
            var Element1 = desktopSession.FindElementsByName("Preliminar");
            //BOTON IMPRIMIR
            Thread.Sleep(500);
            desktopSession.Mouse.MouseMove(Element[11].Coordinates, 200, 17);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Reporte Imprimir", true, file);
            string pdf = @"C:\Reportes\ReportePDF1_" + "Imprimir" + "_" + Hora();
            newFunctions_5.GenerarReportes("Imprimir", file, pdf);

            Thread.Sleep(1000);

            //BOTON PRELIMINAR
            Element1[0].Click();
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Reporte Preliminar", true, file);
            string pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
            newFunctions_5.GenerarReportes("Preliminar", file, pdf1);



        }
    }
}
