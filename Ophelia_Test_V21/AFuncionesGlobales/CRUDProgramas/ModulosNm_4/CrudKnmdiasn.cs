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
    class CrudKnmdiasn : FuncionesVitales
    {
        public static void CRUDKNmDiasn(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
                
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            List<string> data = new List<string>() { crudVars[0], crudVars[1] };
            if (tipo == 1)
            {
                PruebaCRUD.LupaDinamica(desktopSession, data);
                ElementList[8].SendKeys(crudVars[2]);
                ElementList[7].SendKeys(crudVars[3]);
                ElementList[6].SendKeys(crudVars[4]);
            }
            else
            {
                ElementList[6].Clear();
                ElementList[6].SendKeys(crudVars[5]);
            }
        }
    }
}
