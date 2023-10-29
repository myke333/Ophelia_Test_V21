using OpenQA.Selenium.Appium.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_3
{
    class CrudKnmjuzga
    {
        public static void CrudKNmjuzga(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (Tipo == 1)
            {
                ElementList[6].SendKeys(data[0]);
                Thread.Sleep(500);
                ElementList[5].SendKeys(data[1]);
                Thread.Sleep(500);
            }
            else if (Tipo == 2)
            {
                Thread.Sleep(500);
                ElementList[5].Clear();
                Thread.Sleep(500);
                ElementList[5].SendKeys(data[2]);
                Thread.Sleep(500);
            }
        }
    }
}
