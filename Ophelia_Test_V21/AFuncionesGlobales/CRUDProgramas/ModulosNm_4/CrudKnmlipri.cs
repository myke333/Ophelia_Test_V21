using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_4
{
    class CrudKnmlipri : FuncionesVitales
    {
        public static void CrudKNmlipri(WindowsDriver<WindowsElement> desktopSession, string file)
        {
            desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Aceptar").Coordinates);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(10000);
            PruebaCRUD.cerrarBorrar(desktopSession);
            Thread.Sleep(5000);
            newFunctions_4.ScreenshotNuevo("Proceso Ejecutado", true, file);
            PruebaCRUD.cerrarBorrar(desktopSession);
            Thread.Sleep(2000);
            try
            {
                string pdf = @"C:\Reportes\ReportePdf_" + "Errores" + "_" + Hora();
                desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Errores").Coordinates);
                desktopSession.Mouse.Click(null);
                newFunctions_4.ScreenshotNuevo("Errores arrojados por el proceso", true, file);
                Thread.Sleep(500);
                newFunctions_5.GenerarReportes("Errores", file, pdf);
            }
            catch
            {

            }
        }
    }
}
