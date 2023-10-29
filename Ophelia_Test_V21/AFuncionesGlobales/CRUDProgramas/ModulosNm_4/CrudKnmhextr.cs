using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_4
{
    class CrudKnmhextr : PruebaCRUD
    {
        public static void CRUDKnmhextr(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data)
        {
            List<string> errorMessages = new List<string>();

            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList2 = desktopSession.FindElementsByClassName("TDBComboBox");
            if (tipo == 1)
            {
                LupaDinamica2(desktopSession, data);
                ElementList[6].SendKeys(data[1]);
                ElementList[4].SendKeys(data[2]);
                ElementList[5].SendKeys(data[3]);
                ElementList[3].SendKeys(data[4]);
                ElementList[2].SendKeys(data[5]);
                ElementList2[0].SendKeys(data[6]);
            }
            else
            {
                ElementList[2].Clear();
                ElementList[2].SendKeys(data[7]);
            }

        }
    }
}
