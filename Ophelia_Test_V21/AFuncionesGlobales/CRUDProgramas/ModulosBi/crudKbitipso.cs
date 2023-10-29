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


namespace Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium
{
	class crudKbitipso : FuncionesVitales
    {
        public static void CRUDKBiTipso(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            List<string> errorMessages = new List<string>();

            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            int cont = 0;
            if (tipo == 1)
            {
                ElementList[3].SendKeys(crudVars[0]);
                ElementList[2].SendKeys(crudVars[1]);
            }
            else
            {
                ElementList[2].Clear();
                ElementList[2].SendKeys(crudVars[2]);
            }
        }

        public static void CRUDDetalle1KBiTipso(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVarsdet1, string file)
        {
           
            var ElementList = desktopSession.FindElementsByClassName("TDBGridInplaceEdit");
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, 10, 10);
            desktopSession.Mouse.Click(null);
            int cont = 0;
            if (tipo == 1)
            {
                for (int i = 0; i < crudVarsdet1.Count; i++)
                {
                    desktopSession.Keyboard.SendKeys(crudVarsdet1[cont]);
                    cont += 1;
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
                    Thread.Sleep(500);
                }

            }
            else
            {
                ElementList[1].Clear();
                ElementList[1].SendKeys(crudVarsdet1[2]);
            }

        }

        public static void CRUDDetalle2KBiTipso(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVarsdet2, string file)
        {
           
            var ElementList = desktopSession.FindElementsByClassName("TDBGridInplaceEdit");
            int cont = 0;
            if (tipo == 1)
            {
                desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, 10, 10);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                for (int i = 0; i < crudVarsdet2.Count; i++)
                {
                    desktopSession.Keyboard.SendKeys(crudVarsdet2[cont]);
                    cont += 1;
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Right);
                    Thread.Sleep(500);

                }

            }
            else
            {
                ElementList[1].Clear();
                ElementList[1].SendKeys(crudVarsdet2[1]);
            }

        }

    }
}
