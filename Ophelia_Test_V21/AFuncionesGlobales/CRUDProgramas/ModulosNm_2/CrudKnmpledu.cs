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
    class CrudKnmpledu : FuncionesVitales
    {
        public static void CRUDKNmPledu(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {

            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            //Debug.WriteLine("con: " + ElementList.Count);
            //for (int i = 0; i <= ElementList.Count; i++)
            //{
            //    ElementList[i].Click();
            //    Thread.Sleep(500);
            //}
            List<string> data = new List<string>() { crudVars[0], crudVars[1], crudVars[2], crudVars[3] };

            if (tipo == 1)
            {
                PruebaCRUD.LupaDinamica(desktopSession, data);
                ElementList[3].SendKeys(crudVars[4]);
                ElementList[5].SendKeys(crudVars[5]);
            }
            else
            {
                ElementList[5].Clear();
                ElementList[5].SendKeys(crudVars[6]);
            }
        }
    }
}