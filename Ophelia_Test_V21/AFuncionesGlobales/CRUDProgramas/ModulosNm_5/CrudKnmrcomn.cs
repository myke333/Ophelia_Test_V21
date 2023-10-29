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
    class CrudKnmrcomn : FuncionesVitales
    {
        public static void CRUDKNmRcomn(WindowsDriver<WindowsElement> desktopSession, List<string> crudVars, string file)
        {
            var Element = desktopSession.FindElementsByClassName("TEdit");
            Thread.Sleep(500);           
            int var = 0;
            for (int i = Element.Count-1; i >= 0 ; i--)
            {
                Element[i].SendKeys(crudVars[var]);
                var+=1;
            }
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Digitar Datos", true, file);
            Thread.Sleep(500);
        }
        public static void AceptarKNmRcomn(WindowsDriver<WindowsElement> desktopSession, string file)
        {
            var Aceptar = desktopSession.FindElementsByName("Aceptar");
            Thread.Sleep(500);
            Aceptar[0].Click();
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Reporte Comparativo de Nomina", true, file);
            string pdf = @"C:\Reportes\ReportePDF1_" + "Comparativo de Nomina" + "_" + Hora();
            newFunctions_5.GenerarReportes("Comparativo de Nomina", file, pdf);
            Thread.Sleep(500);
        }
    }
}
