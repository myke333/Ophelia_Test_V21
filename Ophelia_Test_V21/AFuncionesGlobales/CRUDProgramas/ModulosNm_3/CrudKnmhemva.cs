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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_3
{
    class CrudKnmhemva : PruebaCRUD
    {
        public static void CRUDKnmhemva(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var Boton = desktopSession.FindElementsByClassName("TDBComboBox");
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            if (tipo == 1)
            {
                LupaDinamica2(desktopSession, data);
                Boton[0].SendKeys(data[1]);
                ElementList[4].SendKeys(data[2]);
            }
            else
            {
                ElementList[4].Clear();
                ElementList[4].SendKeys(data[3]);
            }
        }
    }
}
