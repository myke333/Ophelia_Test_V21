﻿using OpenQA.Selenium.Appium;
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
    class CrudKnmecopr : FuncionesVitales
    {
        public static void CRUDKNmEcopr(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {            
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList2 = desktopSession.FindElementsByName("Seccional");
            List<string> data = new List<string>() { crudVars[0], crudVars[1], crudVars[2], crudVars[3], crudVars[4] };
            if (tipo == 1)
            {
                ElementList2[0].Click();
                PruebaCRUD.LupaDinamica(desktopSession, data);
                ElementList[11].SendKeys(crudVars[5]);
                //ElementList[9].SendKeys(crudVars[1]);
                //ElementList[5].SendKeys(crudVars[2]);
                //ElementList[4].SendKeys(crudVars[3]);
            }
            else
            {
                ElementList[7].Clear();
                ElementList[7].SendKeys(crudVars[4]);
            }
        }
    }
}
