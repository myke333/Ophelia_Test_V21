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
    class CrudKnmbasva : FuncionesVitales
    {
        public static void CRUDKNmBasva(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVars, string file)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");            
            List<string> data = new List<string>() { crudVars[0], crudVars[1], crudVars[2], crudVars[3] };
            if (tipo == 1)
            {
                PruebaCRUD.LupaDinamica(desktopSession, data);
                Thread.Sleep(500);
                ElementList[2].SendKeys(crudVars[4]);
                Thread.Sleep(500);
                ElementList[6].SendKeys(crudVars[5]);
                Thread.Sleep(500);
                ElementList[3].SendKeys(crudVars[6]);
                Thread.Sleep(500);
            }
            else
            {
                ElementList[3].Clear();
                ElementList[3].SendKeys(crudVars[7]);
            }
        }
    }
}
