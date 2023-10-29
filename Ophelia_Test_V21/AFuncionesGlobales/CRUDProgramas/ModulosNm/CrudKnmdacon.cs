using OpenQA.Selenium.Appium.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm
{
    class CrudKnmdacon
    {
        public static void CrudKNmdacon(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (Tipo == 1)
            {
                CrudKnmctrch.LupaDinamica(desktopSession, data);
                ElementList[8].SendKeys(data[6]);
            }
            else if (Tipo == 2)
            {
                ElementList[8].Clear();
                ElementList[8].SendKeys(data[7]);
            }
        }
    }
}
