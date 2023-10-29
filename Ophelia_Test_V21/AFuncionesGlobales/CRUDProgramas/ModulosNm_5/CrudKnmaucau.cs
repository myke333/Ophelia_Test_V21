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
    class CrudKnmaucau
    {
        public static void CrudKNmaucau(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            List<string> lupa = new List<string> { data[0] };
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (Tipo == 1)
            {
                PruebaCRUD.LupaDinamica2(desktopSession, lupa);
                ElementList[2].SendKeys(data[1]);
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
