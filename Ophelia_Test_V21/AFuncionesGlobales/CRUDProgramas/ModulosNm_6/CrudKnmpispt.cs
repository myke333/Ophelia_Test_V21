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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_6
{
    class CrudKnmpispt : FuncionesVitales
    {
        public static void CRUDKNmPispt(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            List<string> data = new List<string>() { crudVars[0], crudVars[1], crudVars[2], crudVars[3] };
            if (tipo == 1)
            {
                PruebaCRUD.LupaDinamica2(desktopSession, data);
                ElementList[16].SendKeys(crudVars[4]);
                ElementList[15].SendKeys(crudVars[5]);
                ElementList[12].SendKeys(crudVars[6]);
                ElementList[9].SendKeys(crudVars[7]);
                ElementList[6].SendKeys(crudVars[8]);
                ElementList[3].SendKeys(crudVars[9]);
                ElementList[2].SendKeys(crudVars[10]);
            }
            else
            {
                ElementList[2].Clear();
                ElementList[2].SendKeys(crudVars[11]);
            }
        }
        public static void CRUDKDetalle1KNmPispt(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TKactusDBgrid");
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, 45, 33);
            desktopSession.Mouse.Click(null);
            var ElementList2 = ElementList[1].FindElementsByClassName("TDBGridInplaceEdit");
            if (ElementList2.Count == 0)
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                ElementList2 = ElementList[1].FindElementsByClassName("TDBGridInplaceEdit");
            }
            if (tipo == 1)
            {
                ElementList2[0].SendKeys(crudVars[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudVars[1]);
            }
            else
            {
                ElementList2[0].Clear();
                ElementList2[0].SendKeys(crudVars[2]);
            }
        }
    }
}
