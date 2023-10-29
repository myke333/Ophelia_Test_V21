using OpenQA.Selenium.Appium.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_3
{
    class CrudKnmgrutr
    {
        public static void CrudKNmgrutr(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (Tipo == 1)
            {
                Thread.Sleep(500);
                ElementList[4].SendKeys(data[0]);
                Thread.Sleep(500);
                ElementList[3].SendKeys(data[1]);
                Thread.Sleep(500);
                ElementList[8].SendKeys(data[2]);
                Thread.Sleep(500);
                ElementList[7].SendKeys(data[3]);
                Thread.Sleep(500);
            }
            else if (Tipo == 2)
            {
                Thread.Sleep(500);
                ElementList[7].Clear();
                Thread.Sleep(500);
                ElementList[7].SendKeys(data[4]);
                Thread.Sleep(500);
            }
        }
    }
}
