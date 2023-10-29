using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSo
{
    class CrudKsoproba : PruebaCRUD
    {
        public static void CRUDKsoproba(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");

            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            if (tipo == 1)
            {
                ElementList[5].SendKeys(data[0]);
                ElementList[2].SendKeys(data[1]);
                ElementList[4].SendKeys(data[2]);
                ElementList[3].SendKeys(data[3]);
            }
            else
            {
                ElementList[3].Clear();
                ElementList[3].SendKeys(data[4]);
            }
        }
    }
}
