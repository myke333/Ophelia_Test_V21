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
    class CrudKnmgdpre : FuncionesVitales
    {
        public static List<string> CrudKNmgdpre(WindowsDriver<WindowsElement> desktopSession, List<string> data, string file, string TipoQbe, string QbeFilter)
        {
            List<string> lupa = new List<string> { data[3] };
            List<string> errors = new List<string>();
            List<string> ErrorValidacion = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TEdit");

            Thread.Sleep(500);
            ElementList[0].Clear();
            Thread.Sleep(500);
            ElementList[0].SendKeys(data[0]);
            Thread.Sleep(500);
            ElementList[4].SendKeys(data[1]);
            Thread.Sleep(500);
            ElementList[3].SendKeys(data[2]);
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Datos Seleccionados", true, file);
            Thread.Sleep(500);
            PruebaCRUD.LupaDinamica2(desktopSession, lupa);

            //QBE
            errors = FuncionesGlobales.QBEQry(desktopSession, TipoQbe, QbeFilter, file);
            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

            desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Preliminar").Coordinates);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Iniciar el Proceso", true, file);
            Thread.Sleep(5000);
            PruebaCRUD.cerrarBorrar(desktopSession);
            Thread.Sleep(5000);
            newFunctions_4.ScreenshotNuevo("Proceso Terminado", true, file);
            PruebaCRUD.cerrarBorrar(desktopSession);
            //try
            //{
            //    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Errores").Coordinates);
            //    desktopSession.Mouse.Click(null);
            //    Thread.Sleep(500);
            //    newFunctions_4.ScreenshotNuevo("Reporte de Errores", true, file);
            //    string ruta = @"C:\Reportes\ReportePDF1_" + "Errores" + "_" + Hora();
            //    errors = newFunctions_5.GenerarReportes("Errores", file, ruta);
            //    if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
            //}
            //catch
            //{

            //}

            return ErrorValidacion;
        }
    }
}
