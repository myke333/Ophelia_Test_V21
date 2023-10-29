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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSo
{
    class CrudKsomatrl : FuncionesVitales
    {
        public static void CRUDKSoMatrl(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (tipo == 1)
            {
                ElementList[5].SendKeys(crudVars[0]);
            }
            else
            {
                ElementList[5].Clear();
                ElementList[5].SendKeys(crudVars[1]);
            }
        }

        public static void CRUDDetalle1KSoMatrl(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVarsdet1, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            desktopSession.Mouse.MouseMove(ElementList[4].Coordinates, 110, 30);
            desktopSession.Mouse.Click(null);
            var ElementList2 = ElementList[4].FindElementsByClassName("TDBGridInplaceEdit");
            if (tipo == 1)
            {
                ElementList2[0].SendKeys(crudVarsdet1[0]);
            }
            else
            {
                ElementList2[0].Clear();
                ElementList2[0].SendKeys(crudVarsdet1[1]);
            }
        }

        public static void CRUDDetalle2KSoMatrl(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVarsdet2, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            desktopSession.Mouse.MouseMove(ElementList[3].Coordinates, 110, 30);
            desktopSession.Mouse.Click(null);
            var ElementList2 = ElementList[3].FindElementsByClassName("TDBGridInplaceEdit");
            if (tipo == 1)
            {
                ElementList2[0].SendKeys(crudVarsdet2[0]);
            }
            else
            {
                ElementList2[0].Clear();
                ElementList2[0].SendKeys(crudVarsdet2[1]);
            }
        }

        public static void CRUDDetalle3KSoMatrl(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVarsdet3, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            if (tipo == 1)
            {

                desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, 200, 35);
                desktopSession.Mouse.Click(null);
                var ElementList2 = ElementList[1].FindElementsByClassName("TDBGridInplaceEdit");
                ElementList2[0].SendKeys(crudVarsdet3[0]);
            }
            else
            {
                var ElementList2 = ElementList[1].FindElementsByClassName("TDBGridInplaceEdit");
                ElementList2[0].Clear();
                ElementList2[0].SendKeys(crudVarsdet3[1]);
            }
        }

        public static void CRUDDetalle4KSoMatrl(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVarsdet4, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            desktopSession.Mouse.MouseMove(ElementList[2].Coordinates, 10, 30);
            desktopSession.Mouse.Click(null);
            var ElementList2 = ElementList[2].FindElementsByClassName("TDBGridInplaceEdit");
            if (tipo == 1)
            {
                ElementList2[0].SendKeys(crudVarsdet4[0]);
            }
            else
            {
                ElementList2[0].Clear();
                ElementList2[0].SendKeys(crudVarsdet4[1]);
            }
        }

        public static void CRUDDetalle5KSoMatrl(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVarsdet5, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 20, 30);
            desktopSession.Mouse.Click(null);
            var ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
            if (tipo == 1)
            {
                ElementList2[0].SendKeys(crudVarsdet5[0]);
            }
            else
            {
                ElementList2[0].Clear();
                ElementList2[0].SendKeys(crudVarsdet5[1]);
            }
        }

    }
}


