using OpenQA.Selenium.Appium.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosGn
{
    class CrudKgnefluj
    {
        public static void CrudKGnefluj(WindowsDriver<WindowsElement> desktopSession, string motor)
        {
            if(motor == "SQL")
            {
                desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Registro Unico de Inscripción en Carrera").Coordinates);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
            }
            else
            {
                desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Cambio de Turno").Coordinates);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
            }
        }
    }
}
