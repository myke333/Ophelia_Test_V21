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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_4
{
    class CrudKnmmceti : FuncionesVitales
    {
        public static void CRUDKNmMceti(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");                       

            if (tipo == 1)
            {
                List<string> data = new List<string>() { crudVars[0], crudVars[1] };
                PruebaCRUD.LupaDinamica2(desktopSession, data);
            }
            else
            {
                ElementList[4].Clear();
                ElementList[4].SendKeys(crudVars[2]);
            }
        }
        public static void CRUDDetalle1KNmMceti(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVarsdet1, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 45, 35);
            desktopSession.Mouse.Click(null);
            var ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
            if (ElementList2.Count == 0)
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
            }
            if (tipo == 1)
            {
                ElementList2[0].SendKeys(crudVarsdet1[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudVarsdet1[1]);
            }
            else
            {
                ElementList2[0].Clear();
                ElementList2[0].SendKeys(crudVarsdet1[2]);
            }
        }
    }
}
