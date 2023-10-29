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
    class CrudKgninfno : PruebaCRUD
    {
        public static void CRUDKgninfno(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data)
        {
            List<string> errorMessages = new List<string>();

            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList2 = desktopSession.FindElementsByClassName("TDBMemo");

            if (tipo == 1)
            {
                LupaDinamica2(desktopSession, data);
                ElementList[7].SendKeys(data[1]);
                ElementList2[0].SendKeys(data[2]);
            }
            else
            {
                ElementList[7].Clear();
                ElementList[7].SendKeys(data[3]);
            }

        }
    }
}
