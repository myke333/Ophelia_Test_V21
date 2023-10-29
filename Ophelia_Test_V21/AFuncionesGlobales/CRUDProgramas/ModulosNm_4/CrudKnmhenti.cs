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
    class CrudKnmhenti : PruebaCRUD
    {
        public static void CRUDKnmhenti(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data)
        {
            List<string> errorMessages = new List<string>();

            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList2 = desktopSession.FindElementsByClassName("TDBCmbCodigoBox");
            if (tipo == 1)
            {
                ElementList2[0].SendKeys(data[0]);
                ElementList[6].SendKeys(data[1]);
                ElementList[5].SendKeys(data[2]);
                ElementList[2].SendKeys(data[3]);
            }
            else
            {
                ElementList[5].Clear();
                ElementList[5].SendKeys(data[4]);
            }

        }
    }
}
