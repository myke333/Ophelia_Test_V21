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
    class CrudKnmracom : FuncionesVitales
    {
        public static void CRUDKNmRacom(WindowsDriver<WindowsElement> desktopSession, string file)
        {
            for (int i = 1 ; i >= 0; i--)
            {
                var Aceptar = desktopSession.FindElementsByName("Aceptar");
                WindowsDriver<WindowsElement> rootSession = null;
                Thread.Sleep(500);
                Aceptar[0].Click();
                Thread.Sleep(500);
                rootSession = PruebaCRUD.RootSession();
                rootSession = PruebaCRUD.ReloadSession(rootSession, "TFrmImpRep");
                var Element = rootSession.FindElementsByClassName("TGroupButton");
                Aceptar = rootSession.FindElementsByName("Aceptar");

                Thread.Sleep(500);
                newFunctions_4.ScreenshotNuevo("Opciones Imprimir Reporte", true, file);
                Thread.Sleep(500);
                Element[i].Click();
                Thread.Sleep(500);
                Aceptar[0].Click();

                rootSession = PruebaCRUD.RootSession();
                rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
                var allFrame = rootSession.FindElementsByClassName("IEFrame");
                Thread.Sleep(500);
                new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) + 80, (allFrame[0].Size.Height / 2) + 80).DoubleClick().Perform();
                Thread.Sleep(500);
                new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) + 150, (allFrame[0].Size.Height / 2) + 80).DoubleClick().Perform();
                Thread.Sleep(1000);
                newFunctions_4.ScreenshotNuevo("Proceso Finalizado", true, file);
                Thread.Sleep(500);
                new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2) + 120, (allFrame[0].Size.Height / 2) + 80).DoubleClick().Perform();
                Thread.Sleep(500);
                newFunctions_4.ScreenshotNuevo("Reporte Comparativo Nomina", true, file);
                string pdf = @"C:\Reportes\ReportePDF1_" + "Comparativo Nomina" + "_" + Hora();
                newFunctions_5.GenerarReportes("Comparativo Nomina", file, pdf);
            }
            
        }
    }
}
