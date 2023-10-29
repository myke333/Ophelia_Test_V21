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
    class CrudKnmprecc
    {
        public static void CrudKNmprecc(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            if (Tipo == 1)
            {
                PruebaCRUD.LupaDinamica2(desktopSession, data);
                Thread.Sleep(2000);
                var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
                var ElementList1 = desktopSession.FindElementsByClassName("TSpinEdit");
                ElementList1[0].Clear();
                ElementList1[0].SendKeys(data[3]);
                Thread.Sleep(500);
                ElementList[2].SendKeys(data[4]);
                Thread.Sleep(500);
                ElementList[9].SendKeys(data[5]);
            }
            else if (Tipo == 2)
            {
                var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
                ElementList[9].Clear();
                ElementList[9].SendKeys(data[6]);
            }
        }
    }
}
