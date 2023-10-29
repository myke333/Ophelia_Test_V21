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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBi
{
    class CrudKbimaetf : FuncionesVitales
    {
        public static void CRUDKBiMaetf(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList1 = desktopSession.FindElementsByClassName("TDBMemo");
            if (tipo == 1)
            {
                ElementList1[0].SendKeys(crudVars[2]);
                ElementList[3].SendKeys(crudVars[1]);
                ElementList[4].SendKeys(crudVars[0]);
            }
            else
            {
                ElementList1[0].Clear();
                ElementList1[0].SendKeys(crudVars[3]);
            }
        }

    }
}
