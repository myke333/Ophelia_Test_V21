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
    class CrudKnmdeman
    {
        public static void CrudKNmdeman(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (Tipo == 1)
            {
                PruebaCRUD.LupaDinamica(desktopSession, data);
                ElementList[6].SendKeys(data[2]);
                ElementList[5].SendKeys(data[3]);
                ElementList[24].SendKeys(data[4]);
                ElementList[23].SendKeys(data[5]);
                ElementList[22].SendKeys(data[6]);
            }
            else if (Tipo == 2)
            {
                ElementList[6].Clear();
                ElementList[6].SendKeys(data[7]);
            }
        }
    }
}
