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
    class CrudKnmicimf : FuncionesVitales
    {
        public static void CRUDKNmIcimf(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {

            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList2 = desktopSession.FindElementsByClassName("TDBLookupComboBox");
            List<string> data = new List<string>() { crudVars[1], crudVars[2], crudVars[3], crudVars[4], crudVars[5] };
            if (tipo == 1)
            {
                ElementList2[0].Click();
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                desktopSession.Keyboard.SendKeys(crudVars[0]);
                PruebaCRUD.LupaDinamica2(desktopSession, data);
            }
            else
            {
                ElementList[6].Clear();
                ElementList[6].SendKeys(crudVars[6]);
            }
        }
    }
}
