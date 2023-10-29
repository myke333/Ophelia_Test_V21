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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_6
{
    class CrudKnmhdeco : FuncionesVitales
    {
        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            if (tipo == 1)
            {
                desktopSession.FindElementByName(crudVars[3]).Click();
                Thread.Sleep(500);
                desktopSession.FindElementByName(crudVars[4]).Click();
                Thread.Sleep(500);
                PruebaCRUD.LupaDinamica(desktopSession, crudVars);
                Thread.Sleep(500);
                var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
                ElementList[6].SendKeys(crudVars[5]);
                ElementList[7].SendKeys(crudVars[6]);

            }
            else
            {
                var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
                ElementList[7].Clear();
                ElementList[7].SendKeys(crudVars[7]);
            }
        }
    }
}
