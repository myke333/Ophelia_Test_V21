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
    class CrudKnmacadm : FuncionesVitales
    {
        public static void CrudKNmacadm(WindowsDriver<WindowsElement> desktopSession, List<string> data, string file)
        {
            string fecha = "09/02/2021";
            var ElementList = desktopSession.FindElementsByClassName("TEdit");
            PruebaCRUD.LupaDinamica2(desktopSession, data);
            Thread.Sleep(500);
            ElementList[4].SendKeys(fecha);
            Thread.Sleep(500);
            desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Aceptar").Coordinates);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(5000);
            string pdf = @"C:\Reportes\ReportePdf_" + "Errores" + "_" + Hora();
            newFunctions_5.GenerarReportes("Errores", file, pdf);
        }
    }
}
