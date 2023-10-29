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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_3
{
    class CrudKnmrecpr : FuncionesVitales
    {
        public static void CRUDKNmRecpr(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            //for (int i = 0; i < ElementList.Count; i++)
            //{
            //    ElementList[i].Click();
            //    Thread.Sleep(500);
            //}
            List<string> data = new List<string>() { crudVars[0], crudVars[1]};
            if (tipo == 1)
            {

                PruebaCRUD.LupaDinamica(desktopSession, data);
            }
            else
            {
                ElementList[4].Clear();
                ElementList[4].SendKeys(crudVars[2]);
            }
        }
    }
    }
