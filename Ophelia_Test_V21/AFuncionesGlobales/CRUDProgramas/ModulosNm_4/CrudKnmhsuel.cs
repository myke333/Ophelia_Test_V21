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
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_4
{
    class CrudKnmhsuel : FuncionesVitales
    {
        public static void CRUDKNmHsuel(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (tipo == 1)
            {
                List<string> data = new List<string>() { crudVars[0] };
                PruebaCRUD.LupaDinamica(desktopSession, data);
                ElementList[15].SendKeys(crudVars[1]);
                ElementList[8].SendKeys(crudVars[2]);                
            }
            else
            {
                ElementList[15].Clear();
                ElementList[15].SendKeys(crudVars[3]);
            }
        }
    }
}
