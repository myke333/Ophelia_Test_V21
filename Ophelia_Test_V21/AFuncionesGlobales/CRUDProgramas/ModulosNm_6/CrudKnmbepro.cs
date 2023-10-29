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
    class CrudKnmbepro : FuncionesVitales
    {
        public static void CRUDKNmBepro(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {            
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            Thread.Sleep(500);
            if (tipo == 1)
            {
                ElementList[13].SendKeys(crudVars[0]);
                Thread.Sleep(500);
                ElementList[12].SendKeys(crudVars[1]);
                Thread.Sleep(500);
            }
            else
            {
                ElementList[12].Clear();
                ElementList[12].SendKeys(crudVars[2]);
                Thread.Sleep(500);
            }
        }
    }
}
