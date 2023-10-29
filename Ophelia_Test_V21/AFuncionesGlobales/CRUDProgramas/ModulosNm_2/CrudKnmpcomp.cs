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
    class CrudKnmpcomp
    {
        public static void CrudKNmpcomp(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            List<string> data1 = new List<string> { data[0] };
            List<string> data2 = new List<string> { data[1] };
            if (Tipo == 1)
            {
                desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("E - Cantidad de Horas Extras").Coordinates);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                PruebaCRUD.LupaDinamica2(desktopSession, data1);
                Thread.Sleep(500);
            }
            else if (Tipo == 2)
            {
                PruebaCRUD.LupaDinamica2(desktopSession, data2);
                Thread.Sleep(500);
            }
        }
    }
}
