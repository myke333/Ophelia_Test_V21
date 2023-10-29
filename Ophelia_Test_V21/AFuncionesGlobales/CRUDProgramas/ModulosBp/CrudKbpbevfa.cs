using OpenQA.Selenium.Appium.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBp
{
    class CrudKbpbevfa
    {
        public static void CrudKBpbevfa(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (Tipo == 1)
            {
                ElementList[10].SendKeys(data[0]);
                Thread.Sleep(500);
                ElementList[9].SendKeys(data[1]);
                Thread.Sleep(500);
                ElementList[5].SendKeys(data[2]);
                Thread.Sleep(500);
                ElementList[2].SendKeys(data[3]);
                Thread.Sleep(500);
            }
            else if (Tipo == 2)
            {
                ElementList[2].Clear();
                Thread.Sleep(500);
                ElementList[2].SendKeys(data[4]);
            }
        }
    }
}
