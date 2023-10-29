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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosGn
{
    class CrudKgnnotas : PruebaCRUD
    {
        public static void CRUDKgnnotas(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data)
        {
            List<string> errorMessages = new List<string>();

            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList2 = desktopSession.FindElementsByClassName("TGroupButton");
            var ElementList3 = desktopSession.FindElementsByClassName("TDBCheckBox");
            if (tipo == 1)
            {
                LupaDinamica2(desktopSession, data);
                ElementList[10].SendKeys(data[2]);
                ElementList[11].SendKeys(data[3]);
                ElementList2[4].Click();
                ElementList3[0].Click();
                ElementList[5].SendKeys(data[4]);
            }
            else
            {
                ElementList[5].Clear();
                ElementList[5].SendKeys(data[5]);
            }

        }
    }
}
