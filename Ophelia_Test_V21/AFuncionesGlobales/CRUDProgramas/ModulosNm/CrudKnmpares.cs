using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm
{
    class CrudKnmpares : PruebaCRUD
    {
        public static void CRUDKnmpares(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            if (tipo == 1)
            {
                ElementList[9].SendKeys(data[0]);
                ElementList[8].SendKeys(data[1]);
            }
            else
            {
                ElementList[8].Clear();
                ElementList[8].SendKeys(data[2]);
            }
        }
    }
}
