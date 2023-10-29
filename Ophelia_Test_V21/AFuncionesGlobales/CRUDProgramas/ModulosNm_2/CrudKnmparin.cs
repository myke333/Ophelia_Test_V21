using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_2
{
    class CrudKnmparin
    {
        public static void CrudKNmparin(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data, string edit)
        {
            if (Tipo == 1)
            {
                PruebaCRUD.LupaDinamica2(desktopSession, data);
                desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Tomar Nombre de la Empresa de Parámetros de Seguridad Social").Coordinates);
                Thread.Sleep(500);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Planilla independiente por Empleado").Coordinates);
                Thread.Sleep(500);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Terceros").Coordinates);
                Thread.Sleep(500);
                desktopSession.Mouse.Click(null);
            }
            else if (Tipo == 2)
            {
                var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
                ElementList[3].Clear();
                ElementList[3].SendKeys(edit);
            }
        }
    }
}
