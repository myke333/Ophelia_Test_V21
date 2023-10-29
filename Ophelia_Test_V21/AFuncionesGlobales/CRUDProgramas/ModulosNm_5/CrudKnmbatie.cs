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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_5
{
    class CrudKnmbatie : FuncionesVitales
    {
        public static void CRUDKNmBatie(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (tipo == 1)
            {
                Thread.Sleep(500);
                ElementList[3].Clear();
                Thread.Sleep(500);
                ElementList[3].SendKeys(crudVars[0]);
                Thread.Sleep(500);
                ElementList[2].SendKeys(crudVars[1]);
                Thread.Sleep(500);
            }
            else
            {
                ElementList[2].Clear();
                ElementList[2].SendKeys(crudVars[2]);
            }
        }
        public static void DetalleKNmBatie(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars1, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TGroupBox");
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 40, 45);
            desktopSession.Mouse.Click(null);
            var ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
            if (ElementList2.Count == 0)
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                ElementList2 = ElementList[0].FindElementsByClassName("TDBGridInplaceEdit");
            }
            if (tipo == 1)
            {
                Thread.Sleep(500);
                ElementList2[0].SendKeys(crudVars1[0]);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudVars1[1]);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(crudVars1[2]);
            }
            else
            {
                ElementList2[0].Clear();
                ElementList2[0].SendKeys(crudVars1[3]);
            }
        }
    }
}
