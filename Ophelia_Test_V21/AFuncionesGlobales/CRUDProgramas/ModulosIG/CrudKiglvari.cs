using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosIG
{
    class CrudKiglvari
    {
        public static List<string> CrudKIglvari(WindowsDriver<WindowsElement> desktopSession, List<string> data, string file, string ruta)
        {
            List<string> lupas = new List<string> { data[1], data[2] };
            List<string> errors = new List<string> { };
            List<string> ErrorValidacion = new List<string> { };
            var ElementList = desktopSession.FindElementsByClassName("TEdit");
            ElementList[4].SendKeys(data[0]);
            Thread.Sleep(500);
            PruebaCRUD.LupaDinamica2(desktopSession, lupas);
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Datos Utilizados", true, file);
            Thread.Sleep(500);
            desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Aceptar").Coordinates);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Proceso Iniciado", true, file);
            PruebaCRUD.cerrarBorrar(desktopSession);
            Thread.Sleep(1000);
            newFunctions_4.ScreenshotNuevo("Proceso Terminado", true, file);
            PruebaCRUD.cerrarBorrar(desktopSession);
            Thread.Sleep(1000);
            desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Errores").Coordinates);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1000);
            errors = newFunctions_5.GenerarReportes("Errores", file, ruta);
            if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }

            return errors;
        }
    }
}
