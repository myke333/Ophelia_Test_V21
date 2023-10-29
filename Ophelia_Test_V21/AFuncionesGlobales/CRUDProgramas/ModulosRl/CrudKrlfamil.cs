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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosRl
{
    class CrudKrlfamil : FuncionesVitales
    {
        public static void CRUDKRlFamil(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList2 = desktopSession.FindElementsByClassName("TGroupButton");
            var ElementList3 = desktopSession.FindElementsByClassName("TEdit");
            Debug.WriteLine(""+ElementList2.Count);
            //for (int i = 0; i < ElementList.Count; i++)
            //{
            //    desktopSession.Mouse.MouseMove(ElementList[i].Coordinates, 10, 10);
            //    desktopSession.Mouse.ContextClick(null);
            //    Thread.Sleep(500);
            //}

            List<string> data = new List<string>() { crudVars[0] };
            if (tipo == 1)
            {
                PruebaCRUD.LupaDinamica(desktopSession, data);
                ElementList[2].SendKeys(crudVars[1]);
                ElementList3[3].SendKeys(crudVars[2]);
                ElementList3[1].SendKeys(crudVars[3]);
                ElementList2[15].Click();

            }
            else
            {
                ElementList3[1].Clear();
                ElementList3[1].SendKeys(crudVars[4]);
            }
        }
    }
}
