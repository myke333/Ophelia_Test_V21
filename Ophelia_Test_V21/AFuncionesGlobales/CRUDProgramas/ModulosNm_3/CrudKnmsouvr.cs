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
    class CrudKnmsouvr
    {
        public static void CrudKNmsouvr(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            List<string> lupa = new List<string> { data[0], data[1], data[2] };
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList1 = desktopSession.FindElementsByClassName("TDBMemo");
            if (Tipo == 1)
            {
                PruebaCRUD.LupaDinamica(desktopSession, lupa);
                Thread.Sleep(500);
                ElementList[13].SendKeys(data[3]);
                Thread.Sleep(500);
                ElementList[11].SendKeys(data[4]);
                Thread.Sleep(500);
                ElementList[10].SendKeys(data[5]);
                Thread.Sleep(500);
                ElementList[8].SendKeys(data[6]);
                Thread.Sleep(500);
                ElementList1[0].SendKeys(data[7]);
                Thread.Sleep(500);
            }
            else if (Tipo == 2)
            {
                Thread.Sleep(500);
                ElementList1[0].Clear();
                Thread.Sleep(500);
                ElementList1[0].SendKeys(data[8]);
                Thread.Sleep(500);
            }
        }
    }
}
