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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_5
{
    class CrudKnmbasir : FuncionesVitales
    {
        public static void CRUDKNmBasir(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList2 = desktopSession.FindElementsByName("No");
            List<string> data = new List<string>() { crudVars[0] };
            Thread.Sleep(500);
            if (tipo == 1)
            {
                PruebaCRUD.LupaDinamica2(desktopSession, data);
                Thread.Sleep(500);
                ElementList2[0].Click();
                Thread.Sleep(500);
                ElementList[33].SendKeys(crudVars[1]);
                for (int i = 6; i <= 27; i++)
                {
                    ElementList[i].SendKeys(crudVars[2]);
                    Thread.Sleep(500);
                }                
                ElementList[32].SendKeys(crudVars[2]);
                Thread.Sleep(500);
                ElementList[31].SendKeys(crudVars[3]);
            }
            else
            {
                ElementList[31].Clear();
                ElementList[31].SendKeys(crudVars[4]);
            }
        }
    }
}
