using OpenQA.Selenium.Appium.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm
{
    class CrudKnmdcado
    {
        public static void CrudKNmdcado(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (Tipo == 1)
            {
                ElementList[6].SendKeys(data[0]);
                ElementList[2].SendKeys(data[1]);
                ElementList[5].SendKeys(data[2]);
                ElementList[4].SendKeys(data[3]);
                ElementList[3].SendKeys(data[4]);
            }
            else if (Tipo == 2)
            {
                ElementList[2].Clear();
                ElementList[2].SendKeys(data[5]);
            }
        }
    }
}
