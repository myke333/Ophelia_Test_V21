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
    class CrudKactipca
    {
        public static void CrudKActipca(WindowsDriver<WindowsElement> desktopSession, int Tipo, string codigo, string abreviatura, string descrip, string edit)
        {
            if (Tipo == 1)
            {
                var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
                ElementList[4].SendKeys(codigo);
                ElementList[2].SendKeys(abreviatura);
                ElementList[3].SendKeys(descrip);
            }
            else if (Tipo == 2)
            {
                var ElementList = desktopSession.FindElementsByClassName("TDBEdit");

                ElementList[3].Clear();
                ElementList[3].SendKeys(edit);
            }
        }
    }
}
