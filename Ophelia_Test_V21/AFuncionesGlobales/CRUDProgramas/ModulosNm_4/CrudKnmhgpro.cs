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
    class CrudKnmhgpro : PruebaCRUD
    {
        public static void CRUDKnmhgpro(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data)
        {
            List<string> errorMessages = new List<string>();

            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (tipo == 1)
            {
                LupaDinamica2(desktopSession, data);
                ElementList[2].SendKeys(data[3]);
            }
            else
            {
                ElementList[2].Clear();
                ElementList[2].SendKeys(data[4]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            }

        }
        
    }
}
