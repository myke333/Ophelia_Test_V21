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
    class CrudKnmestmi : FuncionesVitales
    {
        public static void CRUDKNmEstmi(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList2 = desktopSession.FindElementsByClassName("TDBMemo");
            if (tipo == 1)
            {                
                ElementList[3].SendKeys(crudVars[0]);
                ElementList[2].SendKeys(crudVars[1]);
                ElementList2[0].SendKeys(crudVars[2]);
            }
            else
            {
                ElementList2[0].Clear();
                ElementList2[0].SendKeys(crudVars[3]);
            }
        }
    }
}