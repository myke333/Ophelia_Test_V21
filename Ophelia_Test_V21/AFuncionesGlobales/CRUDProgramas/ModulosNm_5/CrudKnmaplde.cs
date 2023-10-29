using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_5
{
    class CrudKnmaplde
    {
        public static void CrudKNmaplde(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            List<string> lupa = new List<string> { data[0], data[1], data[2] };
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (Tipo == 1)
            {
                PruebaCRUD.LupaDinamica(desktopSession, lupa);
                Thread.Sleep(500);
                ElementList[7].SendKeys(data[3]);
                Thread.Sleep(500);
            }
            else if (Tipo == 2)
            {
                ElementList[7].Clear();
                Thread.Sleep(500);
                ElementList[7].SendKeys(data[4]);
                Thread.Sleep(500);
            }
        }
    }
}
