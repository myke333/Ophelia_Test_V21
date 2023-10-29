using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_5
{
    class CrudKnmliqre : FuncionesVitales
    {
        public static List<string> CrudKNmliqre(WindowsDriver<WindowsElement> desktopSession, List<string> data, string file)
        {
            List<string> errors = new List<string>();
            PruebaCRUD.LupaDinamica(desktopSession, data);
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Datos Seleccionados", true, file);
            Thread.Sleep(500);
            desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Simulación").Coordinates);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(500);
            desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Aceptar").Coordinates);
            desktopSession.Mouse.Click(null);
            newFunctions_4.ScreenshotNuevo("Iniciar el Proceso", true, file);
            Thread.Sleep(1000);
            PruebaCRUD.cerrarBorrar(desktopSession);
            Thread.Sleep(2000);
            PruebaCRUD.cerrarBorrar(desktopSession);
            Thread.Sleep(30000);
            newFunctions_4.ScreenshotNuevo("Proceso Terminado", true, file);
            PruebaCRUD.cerrarBorrar(desktopSession);
            Thread.Sleep(500);
            try
            {
                desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Errores").Coordinates);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                newFunctions_4.ScreenshotNuevo("Reporte de Errores", true, file);
                string ruta = @"C:\Reportes\ReportePDF1_" + "Errores" + "_" + Hora();
                errors = newFunctions_5.GenerarReportes("Errores", file, ruta);
            }
            catch
            {

            }
            
            return errors;
        }
    }
}
