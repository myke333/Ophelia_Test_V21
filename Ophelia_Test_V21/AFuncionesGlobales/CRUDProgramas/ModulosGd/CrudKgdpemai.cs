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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosGd
{
    class CrudKgdpemai : FuncionesVitales
    {
        public static void CRUDKGdPemai(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file, string motor)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList2 = desktopSession.FindElementsByClassName("TKsCmbCodigo");
            var ElementList3 = desktopSession.FindElementsByClassName("TDBMemo");

            if (tipo == 1)
            {
                if (motor == "SQL")
                {
                    ElementList2[1].Click();
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                }
                ElementList[2].SendKeys(crudVars[0]);
                ElementList3[0].SendKeys(crudVars[1]);
            }
            else
            {
                ElementList3[0].Clear();
                ElementList3[0].SendKeys(crudVars[2]);
            }
        }
        public static void CRUDDetalle1KGdPemai(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVarsdet1, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");

            if (tipo == 1)
            {
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 30, 35);
                desktopSession.Mouse.Click(null);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudVarsdet1[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudVarsdet1[1]);
            }
            else
            {
                ElementList[0].Clear();
                ElementList[0].SendKeys(crudVarsdet1[2]);
            }
        }
    }
}
