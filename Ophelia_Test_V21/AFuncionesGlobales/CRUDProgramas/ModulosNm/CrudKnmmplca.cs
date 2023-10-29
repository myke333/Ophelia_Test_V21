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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm
{
    class CrudKnmmplca : FuncionesVitales
    {
        public static void CRUDKNmMplca(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {

            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            //for (int i = 0; i <= ElementList.Count; i++)
            //{
            //    ElementList[i].Click();
            //    Thread.Sleep(500);

            //}
            if (tipo == 1)
            {
                ElementList[3].SendKeys(crudVars[0]);
            }
            else
            {
                ElementList[3].Clear();
                ElementList[3].SendKeys(crudVars[1]);
            }
        }

    }
}