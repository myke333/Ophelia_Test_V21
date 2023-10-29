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
    class CrudKgnccost
    {
        public static void CrudKGnccost(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            List<string> lupa = new List<string> { data[4], data[5], data[6] };
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (Tipo == 1)
            {
                Thread.Sleep(500);
                ElementList[13].SendKeys(data[0]);
                Thread.Sleep(500);
                ElementList[12].SendKeys(data[1]);
                Thread.Sleep(500);
                ElementList[11].SendKeys(data[2]);
                Thread.Sleep(500);
                ElementList[10].SendKeys(data[3]);
                Thread.Sleep(1000);
                PruebaCRUD.LupaDinamica(desktopSession, lupa);
                Thread.Sleep(500);
                ElementList[3].SendKeys(data[7]);
                Thread.Sleep(500);
            }
            else if (Tipo == 2)
            {
                ElementList[12].Clear();
                Thread.Sleep(500);
                ElementList[12].SendKeys(data[8]);
                Thread.Sleep(500);
            }
        }
    }
}
