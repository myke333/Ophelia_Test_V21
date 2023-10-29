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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBi
{
    class CrudKbidotca : FuncionesVitales
    {
        public static void CRUDKBiDotca(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList1 = desktopSession.FindElementsByClassName("TDBCheckBox");
            var ElementList2 = desktopSession.FindElementsByClassName("TDBMemo");

            List<string> data = new List<string>() { crudVars[0] };
            

            if (tipo == 1)
            {
                PruebaCRUD.LupaDinamica(desktopSession, data);
                desktopSession.Mouse.MouseMove(ElementList1[5].Coordinates, 4, 7);
                desktopSession.Mouse.Click(null);
                ElementList[7].SendKeys(crudVars[1]);
                ElementList[5].SendKeys(crudVars[2]);
                ElementList2[0].SendKeys(crudVars[3]);
            }
            else
            {
                ElementList2[0].Clear();
                ElementList2[0].SendKeys(crudVars[4]);
            }
        }
    }
}
