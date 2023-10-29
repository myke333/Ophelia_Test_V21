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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_2
{
    class CrudKnmpinde : FuncionesVitales
    {
        public static void CRUDKNmPinde(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (tipo == 1)
            {
                ElementList[5].SendKeys(crudVars[0]);
                ElementList[4].SendKeys(crudVars[1]);
                ElementList[3].SendKeys(crudVars[2]);
                ElementList[2].SendKeys(crudVars[3]);
            }
            else
            {
                ElementList[4].Clear();
                ElementList[4].SendKeys(crudVars[4]);
            }
        }

        public static void CRUDDetalle1KNmPinde(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVarsDet1, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGridInplaceEdit");            
            if (tipo == 1)
            {
                ElementList[1].SendKeys(crudVarsDet1[0]);
                for (int i = 1; i <= 5; i++)
                {
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    desktopSession.Keyboard.SendKeys(crudVarsDet1[i]);
                }
            }
            else
            {
                ElementList[1].Clear();
                ElementList[1].SendKeys(crudVarsDet1[6]);
            }
        }
    }
}