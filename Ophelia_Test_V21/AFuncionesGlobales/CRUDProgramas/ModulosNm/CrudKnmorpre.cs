using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm
{
    class CrudKnmorpre
    {
        public static void CrudKNmorpre(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> lupas)
        {
            if (Tipo == 1)
            {
                var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
                PruebaCRUD.LupaDinamica(desktopSession, lupas);
                Thread.Sleep(500);
                ElementList[11].SendKeys(lupas[6]);
                Thread.Sleep(500);
                ElementList[10].SendKeys(lupas[4]);
                Thread.Sleep(500);
                ElementList[9].SendKeys(lupas[2]);
                Thread.Sleep(500);
                ElementList[8].SendKeys(lupas[6]);
                Thread.Sleep(500);
                ElementList[2].SendKeys(lupas[7]);

                ElementList[12].SendKeys(lupas[5]);
                Thread.Sleep(500);
                ElementList[13].SendKeys(lupas[4]);
                Thread.Sleep(500);
                ElementList[14].SendKeys(lupas[3]);                
                Thread.Sleep(500);
                

            }
            else if (Tipo == 2)
            {
                var ElementList = desktopSession.FindElementsByClassName("TDBEdit");

                ElementList[2].Clear();
                ElementList[2].SendKeys(lupas[8]);
            }
        }
    }
}
