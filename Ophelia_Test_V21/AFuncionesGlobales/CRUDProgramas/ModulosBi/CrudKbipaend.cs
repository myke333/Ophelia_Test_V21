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
    class CrudKbipaend : FuncionesVitales
    {
        public static void CRUDKBiPaend(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file, string motor)
        {
            List<string> data = new List<string>() { crudVars[0] };            
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList2 = desktopSession.FindElementsByName("Centro de Costos en Grupo de Implementos");
            if (tipo == 1)
            {
                PruebaCRUD.LupaDinamica(desktopSession, data);
                ElementList[5].Clear();
                ElementList[5].SendKeys(crudVars[1]);
                ElementList[4].SendKeys(crudVars[2]);
                if (motor=="SQL")
                {
                ElementList[9].SendKeys(crudVars[3]);
                ElementList[8].SendKeys(crudVars[4]);
                ElementList[7].SendKeys(crudVars[5]);
                ElementList[6].SendKeys(crudVars[6]);
                ElementList2[0].Click();
                }
            }
            else
            {
                ElementList[4].Clear();
                if (motor == "SQL")
                {
                    ElementList[4].SendKeys(crudVars[7]);
                }
                if (motor == "ORA")
                {
                    ElementList[4].SendKeys(crudVars[3]);
                }
            }
        }
        public static void CRUDDetalle1KBiPaend(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVarsdet1, string file)
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
                ElementList[0].SendKeys(crudVarsdet1[0]);
            }
            else
            {
                ElementList2[0].Clear();
                ElementList2[0].SendKeys(crudVarsdet1[1]);
            }
        }
    }
    }
