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
    class CrudKnmhreli
    {
        public static void CrudKNmhreli(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (Tipo == 1)
            {
                List<string> lupa = new List<string> { data[0], data[4] };
                PruebaCRUD.LupaDinamica(desktopSession, lupa);
                Thread.Sleep(500);
                ElementList[5].SendKeys(data[1]);
                Thread.Sleep(500);
                ElementList[4].SendKeys(data[2]);
                Thread.Sleep(500);
                ElementList[11].SendKeys(data[3]);
                Thread.Sleep(500);
            }
            else if (Tipo == 2)
            {
                Thread.Sleep(500);
                ElementList[11].Clear();
                Thread.Sleep(500);
                ElementList[11].SendKeys(data[5]);
                Thread.Sleep(500);
            }
        }
    }
}
