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
    class CrudKrlempdo : FuncionesVitales
    {
        public static void CRUDKRlEmpdo(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList2 = desktopSession.FindElementsByClassName("TDBMemo");
            var ElementList3 = desktopSession.FindElementsByClassName("TComboBox");

            //for (int i = 0; i < ElementList.Count; i++)
            //{
            //    desktopSession.Mouse.MouseMove(ElementList[i].Coordinates, 10, 10);
            //    desktopSession.Mouse.ContextClick(null);
            //    Thread.Sleep(500);
            //}

            List<string> data = new List<string>() { crudVars[0], crudVars[1] };
            if (tipo == 1)
            {
                PruebaCRUD.LupaDinamica(desktopSession, data);
                ElementList[16].SendKeys(crudVars[2]);
                ElementList2[0].SendKeys(crudVars[3]);
                ElementList3[0].Click();
                ElementList3[0].SendKeys("E");
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

            }
            else
            {
                ElementList2[0].Clear();
                ElementList2[0].SendKeys(crudVars[4]);
            }
        }

    }
}

