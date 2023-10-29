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
    class CrudKnmacten
    {
        public static void CrudKNmacten(WindowsDriver<WindowsElement> desktopSession, List<string> data, string file)
        {
            PruebaCRUD.LupaDinamica(desktopSession, data);
            Thread.Sleep(500);
            desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Aceptar").Coordinates);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(10000);
            newFunctions_4.ScreenshotNuevo("Proceso Terminado", true, file);
            PruebaCRUD.cerrarBorrar(desktopSession);
        }
    }
}
