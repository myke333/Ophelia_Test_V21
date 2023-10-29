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
    class CrudKnmpabon
    {
        public static void CrudKNmpabon(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> lupas, string edit)
        {
            if (Tipo == 1)
            {
                PruebaCRUD.LupaDinamica(desktopSession, lupas);
            }
            else if (Tipo == 2)
            {
                var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
                ElementList[3].Clear();
                ElementList[3].SendKeys(edit);
            }
        }
    }
}
