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
    class CrudKnmmreti : FuncionesVitales
    {
        public static void CRUDKNmMreti(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            //for (int i = 0; i <= ElementList.Count; i++)
            //{
            //    ElementList[i].Click();
            //    Thread.Sleep(500);
            //}
            if (tipo == 1)
            {
                ElementList[6].SendKeys(crudVars[0]);
                ElementList[5].SendKeys(crudVars[3]);
                ElementList[4].SendKeys(crudVars[1]);
                ElementList[2].SendKeys(crudVars[2]);
            }
            else
            {
                ElementList[5].Clear();
                ElementList[5].SendKeys(crudVars[4]);
            }
        }

        public static void CRUDDetalle1KNmMreti(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVarsdet1, string file)
        {

            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            //for (int i = 0; i <= ElementList.Count; i++)
            //{
            //    ElementList1[i].Click();
            //    Thread.Sleep(500);
            //}
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 44, 31);
            desktopSession.Mouse.DoubleClick(null);
            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            if (tipo == 1)
            {               
                desktopSession.Keyboard.SendKeys(crudVarsdet1[0]);
            }
            else
            {                
                desktopSession.Keyboard.SendKeys(crudVarsdet1[1]);
            }
        }

    }
}