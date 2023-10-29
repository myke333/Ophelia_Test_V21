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
    class CrudKnmacinr
    {
        public static void CrudKNmacinr(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            List<string> lupa = new List<string> { data[0] };
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (Tipo == 1)
            {
                PruebaCRUD.LupaDinamica(desktopSession, lupa);
                Thread.Sleep(500);
                ElementList[9].SendKeys(data[1]);
                Thread.Sleep(500);
                ElementList[8].SendKeys(data[2]);
                Thread.Sleep(500);
            }
            else if (Tipo == 2)
            {
                ElementList[8].Clear();
                Thread.Sleep(500);
                ElementList[8].SendKeys(data[3]);
                Thread.Sleep(500);
            }
        }
    }
}
