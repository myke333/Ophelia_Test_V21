using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Ophelia_Test_V21.AFuncionesGlobales;
using OpenQA.Selenium;
using System.Drawing;
using System.Windows.Forms;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm
{
    class CrudKnmmpcar : FuncionesVitales
    {
        public static void CRUDKBIMaefe(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {

            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            List<string> data = new List<string>() { crudVars[0] };
            if (tipo == 1)
            {
                PruebaCRUD.LupaDinamica(desktopSession, data);
                ElementList[3].SendKeys(crudVars[1]);
            }
            else
            {
                ElementList[3].Clear();
                ElementList[3].SendKeys(crudVars[2]);
            }
        }

    }
}

