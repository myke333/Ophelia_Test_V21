using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_3
{
    class CrudKnmhconc
    {
        public static void CrudKNmhconc(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (Tipo == 1)
            {
                List<string> lupa = new List<string> { data[0], data[3] };
                PruebaCRUD.LupaDinamica(desktopSession, lupa);
                Thread.Sleep(500);
                ElementList[6].SendKeys(data[1]);
                Thread.Sleep(500);
                ElementList[4].SendKeys(data[2]);
                Thread.Sleep(500);
            }
            else if (Tipo == 2)
            {
                Thread.Sleep(500);
                ElementList[4].Clear();
                Thread.Sleep(500);
                ElementList[4].SendKeys(data[4]);
                Thread.Sleep(500);
            }
        }
    }
}
