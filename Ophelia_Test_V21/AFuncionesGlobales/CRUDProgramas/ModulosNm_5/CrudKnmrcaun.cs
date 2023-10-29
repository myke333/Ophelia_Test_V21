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
    class CrudKnmrcaun : FuncionesVitales
    {       
        public static void CRUDKNmRcaun(WindowsDriver<WindowsElement> desktopSession, List<string> crudVars, string file)
        {
            var Element = desktopSession.FindElementsByClassName("TSpinEdit");
            var Element1 = desktopSession.FindElementsByName("Datos");
            var Element2 = desktopSession.FindElementsByName("Imprimir");
            Thread.Sleep(500);
            Element[0].SendKeys(crudVars[0]);
            Thread.Sleep(500);
            Element1[0].Click();
            Thread.Sleep(500);
            //try
            //{
            //    //BOTON IMPRIMIR
            //    desktopSession.Mouse.MouseMove(Element2[0].Coordinates, 60, 20);
            //    desktopSession.Mouse.Click(null);
            //    Thread.Sleep(500);
            //    newFunctions_4.ScreenshotNuevo("Reporte Imprimir", true, file);
            //    string pdf = @"C:\Reportes\ReportePDF1_" + "Imprimir" + "_" + Hora();
            //    newFunctions_5.GenerarReportes("Imprimir", file, pdf);
            //    Thread.Sleep(500);
            //}
            //catch
            //{
            //    Thread.Sleep(500);
            //}
           
        }
    }
}
