using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm
{
    class CrudKnmcopri : PruebaCRUD
    {
        public static void CRUDKnmcopri(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            var ElementList1 = desktopSession.FindElementsByClassName("TDBEdit");
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            if (tipo == 1)
            {
                LupaDinamica(rootSession, data);
                ElementList1[17].SendKeys(data[2]);
            }
            else
            {
                ElementList1[17].Clear();
                ElementList1[17].SendKeys(data[3]);
            }
        }
    }
}
