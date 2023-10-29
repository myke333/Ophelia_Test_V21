using OpenQA.Selenium.Appium;
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
    class CrudKnmliaps
    {
        public static void CrudKNmliaps(WindowsDriver<WindowsElement> desktopSession, string motor, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TKsCmbCodigo");
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);
            if(motor == "SQL")
            {
                desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("PRUEBA PARAFISCALES NEGATIVOS").Coordinates);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
            }
            else
            {
                desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("ESAP").Coordinates);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
            }
            
            desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Aceptar").Coordinates);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(10000);
            PruebaCRUD.cerrarBorrar(desktopSession);
            Thread.Sleep(5000);
            newFunctions_4.ScreenshotNuevo("Proceso Ejecutado", true, file);
            PruebaCRUD.cerrarBorrar(desktopSession);
        }
    }
}
