using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosGn
{
    class CrudKgnktabl
    {
        public static void CrudKGnktabl(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            List<string> lupa = new List<string> { data[3] };
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (Tipo == 1)
            {
                ElementList[7].SendKeys(data[0]);
                Thread.Sleep(500);
                ElementList[6].SendKeys(data[1]);
                Thread.Sleep(500);
                ElementList[2].SendKeys(data[2]);
                Thread.Sleep(500);
                PruebaCRUD.LupaDinamica2(desktopSession, lupa);
                Thread.Sleep(500);
                ElementList[3].SendKeys(data[4]);
                Thread.Sleep(500);
                desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Usa Sustitución Patronal").Coordinates);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
            }
            else if (Tipo == 2)
            {
                ElementList[6].Clear();
                Thread.Sleep(500);
                ElementList[6].SendKeys(data[5]);
                Thread.Sleep(500);
            }
        }
    }
}
