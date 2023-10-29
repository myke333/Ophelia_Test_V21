using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm
{
    class CrudKnmconar : PruebaCRUD
    {
        public static void CRUDKnmconar(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string motor)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            if (tipo == 1)
            {
                if (motor == "SQL")
                {
                    LupaDinamica(rootSession, data);
                }else
                {
                    ElementList[15].SendKeys(data[2]);
                    ElementList[19].SendKeys(data[0]);
                    ElementList[12].SendKeys(data[1]);
                    ElementList[4].SendKeys(data[1]);
                    ElementList[10].SendKeys(data[2]);
                    ElementList[3].SendKeys(data[2]);
                    ElementList[8].SendKeys(data[3]);
                    ElementList[2].SendKeys(data[3]);
                    ElementList[6].SendKeys(data[4]);
                }
                
                ElementList[14].Clear();
                ElementList[14].SendKeys(data[5]);
            }
            else
            {
                ElementList[6].Clear();
                ElementList[6].SendKeys(data[6]);
            }
        }
    }
}
