using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSo
{
    class CrudKsocodia
    {
        public static void CrudKSocodia(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList1 = desktopSession.FindElementsByClassName("TDBMemo");
            List<string> lupa = new List<string> { data[2] };
            if (Tipo == 1)
            {
                ElementList[4].SendKeys(data[0]);
                Thread.Sleep(500);
                ElementList1[0].SendKeys(data[1]);
                Thread.Sleep(500);
                PruebaCRUD.LupaDinamica2(desktopSession, lupa);
                Thread.Sleep(500);
                ElementList1[1].SendKeys(data[3]);
                Thread.Sleep(500);
            }
            else if (Tipo == 2)
            {
                ElementList1[1].Clear();
                Thread.Sleep(500);
                ElementList1[1].SendKeys(data[4]);
                Thread.Sleep(500);
            }
        }
    }
}
