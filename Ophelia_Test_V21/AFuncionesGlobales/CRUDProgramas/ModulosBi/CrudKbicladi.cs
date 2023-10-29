using OpenQA.Selenium.Appium.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBi
{
    class CrudKbicladi
    {
        public static void CrudKBicladiCabecera(WindowsDriver<WindowsElement> rootSession, int Tipo, string Codigo, string Nombre, string Editar)
        {
            var ElementList = rootSession.FindElementsByClassName("TDBEdit");

            if (Tipo == 1)
            {
                ElementList[3].SendKeys(Codigo);
                ElementList[2].SendKeys(Nombre);
            }
            else if (Tipo == 2)
            {
                ElementList[2].Clear();
                ElementList[2].SendKeys(Editar);
            }
        }

        public static void CrudKBicladiDetalle(WindowsDriver<WindowsElement> rootSession, int Tipo, string Diagnostico, string NomDiag, string Editar)
        {
            var ElementList = rootSession.FindElementsByClassName("TDBGrid");
            rootSession.Mouse.MouseMove(ElementList[0].Coordinates, 66, 26);
            rootSession.Mouse.Click(null);

            if (Tipo == 1)
            {
                ElementList[0].SendKeys(Diagnostico);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                ElementList[0].SendKeys(NomDiag);
            }
            else if (Tipo == 2)
            {
                ElementList[0].Clear();
                ElementList[0].SendKeys(Editar);
            }
        }
    }
}
