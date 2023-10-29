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
    class CrudKacrequi
    {
        public static void CrudKAcrequi(WindowsDriver<WindowsElement> desktopSession, int Tipo, string codigo, string nombre, List<string> Lupa, string edit)
        {
            if (Tipo == 1)
            {
                var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
                ElementList[5].SendKeys(codigo);
                ElementList[4].SendKeys(nombre);
                Thread.Sleep(1000);
                PruebaCRUD.LupaDinamica(desktopSession, Lupa);
            }
            else if (Tipo == 2)
            {
                var ElementList = desktopSession.FindElementsByClassName("TDBEdit");

                ElementList[4].Clear();
                ElementList[4].SendKeys(edit);
            }
        }
    }
}
