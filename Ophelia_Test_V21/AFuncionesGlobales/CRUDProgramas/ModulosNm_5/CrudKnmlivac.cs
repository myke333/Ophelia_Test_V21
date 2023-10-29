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
    class CrudKnmlivac : FuncionesVitales
    {
        public static void CrudKNmlivac(WindowsDriver<WindowsElement> desktopSession, string file)
        {
            desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Simulación Proceso").Coordinates);
            desktopSession.Mouse.Click(null);
            //desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Aceptar").Coordinates);
            //desktopSession.Mouse.Click(null);
            //Thread.Sleep(10000);
            //PruebaCRUD.cerrarBorrar(desktopSession);
            //Thread.Sleep(5000);
            //newFunctions_4.ScreenshotNuevo("Proceso Ejecutado", true, file);
            //PruebaCRUD.cerrarBorrar(desktopSession);
            //Thread.Sleep(2000);
            //try
            //{
            //    string pdf = @"C:\Reportes\ReportePdf_" + "Errores" + "_" + Hora();
            //    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Informe").Coordinates);
            //    desktopSession.Mouse.Click(null);
            //    newFunctions_4.ScreenshotNuevo("Informe del Proceso", true, file);
            //    Thread.Sleep(500);
            //    newFunctions_5.GenerarReportes("Errores", file, pdf);
            //}
            //catch
            //{

            //}
        }
    }
}
