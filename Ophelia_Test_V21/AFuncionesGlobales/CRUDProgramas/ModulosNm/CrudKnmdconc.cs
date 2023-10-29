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
    class CrudKnmdconc
    {
        public static void CrudKNmdconc(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (Tipo == 1)
            {
                PruebaCRUD.LupaDinamica(desktopSession, data);
                ElementList[3].SendKeys(data[1]);
                ElementList[2].SendKeys(data[2]);
            }
            else if (Tipo == 2)
            {
                ElementList[2].Clear();
                ElementList[2].SendKeys(data[3]);
            }
        }
    }
}
