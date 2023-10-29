using OpenQA.Selenium.Appium.Windows;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBi
{
    class CrudKbiunesp
    {
        public static void CrudKBiunesp(WindowsDriver<WindowsElement> rootSession, int Tipo, string Codigo, string Nombre, string Especificacion, string Editar)
        {
            var ElementList = rootSession.FindElementsByClassName("TDBEdit");

            if (Tipo == 1)
            {
                ElementList[4].SendKeys(Codigo);
                ElementList[3].SendKeys(Nombre);
                ElementList[2].SendKeys(Especificacion);
            }
            else if (Tipo == 2)
            {
                ElementList[2].Clear();
                ElementList[2].SendKeys(Editar);
            }
        }
    }
}
