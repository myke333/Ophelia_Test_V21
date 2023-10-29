using OpenQA.Selenium.Appium.Windows;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBi
{
    class CrudKbihsidi
    {
        public static void CrudKBihsidi(WindowsDriver<WindowsElement> rootSession, int Tipo, string Nombre, string Codigo, string Editar)
        {
            var ElementList = rootSession.FindElementsByClassName("TDBEdit");

            if (Tipo == 1)
            {
                ElementList[3].SendKeys(Nombre);
                ElementList[2].SendKeys(Codigo);
            }
            else if (Tipo == 2)
            {
                ElementList[3].Clear();
                ElementList[3].SendKeys(Editar);
            }
        }
    }
}
