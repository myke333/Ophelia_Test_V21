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
    class CrudKnmbecon : FuncionesVitales
    {
        public static void CRUDKNmBecon(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            List<int> dis = new List<int>() { 0, 1 };
            if (tipo == 1)
            {
                List<string> data = new List<string>() { crudVars[0] };
                newFunctions_4.LupaDinamicaDiscriminatoria(desktopSession, data, dis);
                ElementList[4].SendKeys(crudVars[1]);
                Thread.Sleep(500);
            }
            else
            {
                List<string> data = new List<string>() { crudVars[2] };
                newFunctions_4.LupaDinamicaDiscriminatoria(desktopSession, data, dis);
            }
        }
    }
}
