using OpenQA.Selenium.Appium.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosFd
{
    class CrudKfdarcon
    {
        public static void CrudKFdarcon(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList1 = desktopSession.FindElementsByClassName("TDBMemo");
            if (Tipo == 1)
            {
                ElementList[2].SendKeys(data[0]);
                ElementList1[0].SendKeys(data[1]);
            }
            else if (Tipo == 2)
            {
                ElementList1[0].Clear();
                ElementList1[0].SendKeys(data[2]);
            }
        }
    }
}
