using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosAc
{
    class CrudKtiprg
    {
        public static void CrudKTiprg(WindowsDriver<WindowsElement> rootSession, int Tipo, string codigo, string nomReg, string edit)
        {
            if (Tipo == 1)
            {
                var ElementList = rootSession.FindElementsByClassName("TDBEdit");
                ElementList[3].SendKeys(codigo);
                ElementList[2].SendKeys(nomReg);
            }
            else if (Tipo == 2)
            {
                var ElementList = rootSession.FindElementsByClassName("TDBEdit");

                ElementList[2].Clear();
                ElementList[2].SendKeys(edit);
            }
        }
    }
}
