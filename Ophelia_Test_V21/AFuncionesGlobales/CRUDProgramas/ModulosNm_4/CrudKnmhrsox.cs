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
    class CrudKnmhrsox : FuncionesVitales
    {
        public static void CRUDKNmHrsox(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file, string motor)
        {      
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList2 = desktopSession.FindElementsByName("Voluntarios");
            if (tipo == 1)
            {
                if (motor == "SQL")
                {
                    List<string> data = new List<string>() { crudVars[0] };
                    PruebaCRUD.LupaDinamica2(desktopSession, data);
                }                
                ElementList2[0].Click();
                ElementList[6].SendKeys(crudVars[1]);
                ElementList[4].SendKeys(crudVars[2]);
                if (motor == "ORA")
                {
                    ElementList[8].SendKeys(crudVars[0]);
                    ElementList[5].SendKeys(crudVars[4]);
                }
            }
            else
            {
                ElementList[6].Clear();
                ElementList[6].SendKeys(crudVars[3]);
            }
        }
    }
}
