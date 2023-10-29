using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosRl
{
    class CrudKrlmedio : PruebaCRUD
    {
        public static void CRUDKrlmedio(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList2 = desktopSession.FindElementsByClassName("TDBCmbCodigoBox");
            
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            if (tipo == 1)
            {
                ElementList[4].SendKeys(data[0]);
                ElementList2[0].SendKeys(data[1]);
                ElementList[3].SendKeys(data[2]);
            }
            else
            {
                ElementList[3].Clear();
                ElementList[3].SendKeys(data[3]);
            }
        }
    }
}
