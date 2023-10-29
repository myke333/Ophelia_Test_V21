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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBi
{
    class CrudKbimaefe : FuncionesVitales
    {
        public static void CRUDKBIMaefe(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            Debug.WriteLine("contador: " + ElementList);
            List<string> data = new List<string>() { crudVars[0] };
            PruebaCRUD.LupaDinamica(desktopSession, data);
            if (tipo == 1)
            {
                ElementList[4].SendKeys(crudVars[2]);
                ElementList[2].SendKeys(crudVars[1]);
            }
            else
            {
                ElementList[2].Clear();
                ElementList[2].SendKeys(crudVars[3]);
            }
        }

        public static void CRUDDetalle1KBIMaefe(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVarsdet1, string file)
        {
            
            var ElementList = desktopSession.FindElementsByClassName("TDBMemo");
            if (tipo == 1)
            {
                ElementList[1].Clear();
                ElementList[1].SendKeys(crudVarsdet1[0]);
            }
            else
            {
                ElementList[1].Clear();
                ElementList[1].SendKeys(crudVarsdet1[1]);
            }

        }
    }
}
