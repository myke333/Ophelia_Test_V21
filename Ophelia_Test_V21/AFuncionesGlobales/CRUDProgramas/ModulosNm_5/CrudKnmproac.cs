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
    class CrudKnmproac : FuncionesVitales
    {
        public static void CRUDKNmProac(WindowsDriver<WindowsElement> desktopSession, string file)
        {
            var Aceptar = desktopSession.FindElementsByName("Aceptar");
            Aceptar[0].Click();
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Proceso Finalizado", true, file);
            Thread.Sleep(500);
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = PruebaCRUD.RootSession();
            rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");

            var allFrame = rootSession.FindElementsByClassName("IEFrame");
            Thread.Sleep(5000);
            new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) + 110, (allFrame[0].Size.Height / 2) + 90).DoubleClick().Perform();
            Thread.Sleep(5000);
            new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) + 150, (allFrame[0].Size.Height / 2) + 90).DoubleClick().Perform();
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Proceso Finalizado", true, file);
            Thread.Sleep(4500);
            new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) + 110, (allFrame[0].Size.Height / 2) + 90).DoubleClick().Perform();

            //BOTON Reporte
            var Reporte = desktopSession.FindElementsByName("Reporte");
            Thread.Sleep(500);
            Reporte[0].Click();
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Reporte", true, file);
            string pdf = @"C:\Reportes\ReportePDF1_" + "Reporte" + "_" + Hora();
            newFunctions_5.GenerarReportes("Errores", file, pdf);
        }
    }
}
