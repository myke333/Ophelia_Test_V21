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
    class CrudKnmracmc : FuncionesVitales
    {
        public static void CRUDKNmRacmc(WindowsDriver<WindowsElement> desktopSession, string file)
        {
            //BOTON IMPIMIR
            WindowsElement printButton = desktopSession.FindElementByName("Imprimir");

            desktopSession.Mouse.MouseMove(printButton.Coordinates, printButton.Size.Width / 2, printButton.Size.Height / 2);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Reporte Imprimir", true, file);
            string pdf = @"C:\Reportes\ReportePDF1_" + "Imprimir" + "_" + Hora();
            newFunctions_5.GenerarReportes("Errores", file, pdf);
        }
    }
}
