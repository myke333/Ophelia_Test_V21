using OpenQA.Selenium.Appium.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBp
{
    class CrudKbprecev
    {
        public static void CrudKBprecev(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (Tipo == 1)
            {
                ElementList[13].SendKeys(data[0]);
                Thread.Sleep(500);
                ElementList[11].SendKeys(data[1]);
                Thread.Sleep(500);
                ElementList[9].SendKeys(data[2]);
                Thread.Sleep(500);
                ElementList[6].SendKeys(data[3]);
                Thread.Sleep(500);
                ElementList[3].SendKeys(data[4]);
                Thread.Sleep(500);
                ElementList[8].SendKeys(data[5]);
                Thread.Sleep(500);
                ElementList[4].SendKeys(data[6]);
                Thread.Sleep(500);
            }
            else if (Tipo == 2)
            {
                ElementList[4].Clear();
                Thread.Sleep(500);
                ElementList[4].SendKeys(data[7]);
            }
        }
    }
}
