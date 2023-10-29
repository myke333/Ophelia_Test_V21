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
    class CrudKnmbsavi : FuncionesVitales
    {
        public static void CRUDKNmBsavi(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            List<string> data = new List<string>() { };
            if (tipo == 1)
            {
                data = new List<string>() { crudVars[0] };
            }
            else
            {
                data = new List<string>() { crudVars[1] };
            }
            PruebaCRUD.LupaDinamica2(desktopSession, data);
        }
        public static void CRUDDetalleKNmBsavi(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars1, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TGroupBox");
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 40, 50);
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
            }
            else
            {
                ElementList2[0].Clear();
                ElementList2[0].SendKeys(crudVars1[2]);
            }
        }
    }
}
