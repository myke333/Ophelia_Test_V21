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
    class CrudKnmprold : FuncionesVitales
    {
        public static void CRUDKNmProld(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {

            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList2 = desktopSession.FindElementsByClassName("TGroupBox");
            var ElementList3 = ElementList2[0].FindElementsByClassName("TDBCheckBox");
            List<string> data = new List<string>() { crudVars[0], crudVars[1], crudVars[2], crudVars[3], crudVars[4] };
            if (tipo == 1)
            {

                PruebaCRUD.LupaDinamica(desktopSession, data);
                ElementList3[0].Click();
            }
            else
            {
                ElementList[3].Clear();
                ElementList[3].SendKeys(crudVars[5]);
            }
        }
        public static void CRUDDetalleKNmProld(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVarsdet1, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            var ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
            if (tipo == 1)
            {
                ElementList[0].Click();
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 50,11);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudVarsdet1[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudVarsdet1[1]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudVarsdet1[2]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudVarsdet1[3]);
            }
            else
            {
                ElementList[0].Clear();
                ElementList[0].SendKeys(crudVarsdet1[4]);
            }
        }
    }
}
