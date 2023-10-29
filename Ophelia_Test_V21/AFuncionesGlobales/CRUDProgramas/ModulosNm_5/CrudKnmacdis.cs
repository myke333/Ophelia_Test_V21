using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_5
{
    class CrudKnmacdis
    {
        public static void CrudKNmacdis(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (Tipo == 1)
            {
                ElementList[3].SendKeys(data[0]);
                Thread.Sleep(500);
                ElementList[2].SendKeys(data[1]);
                Thread.Sleep(500);
                ElementList[2].SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
            }
            else if (Tipo == 2)
            {
                ElementList[2].Clear();
                Thread.Sleep(500);
                ElementList[2].SendKeys(data[2]);
                Thread.Sleep(500);
            }
        }
    }
}
