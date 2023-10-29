using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm
{
    class CrudKnmconva : PruebaCRUD
    {
        public static void CRUDKnmconva(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            var Pestania = desktopSession.FindElementsByName("Datos Consolidación");
            Thread.Sleep(100);
            Pestania[0].Click(); 
            Thread.Sleep(100);
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            Thread.Sleep(100);
            var ElementList1 = rootSession.FindElementsByClassName("TDBEdit");
            Thread.Sleep(1000);
            if (tipo == 1)
            {
                LupaDinamica(rootSession, data);
                //Debugger.Launch();
                //for (int i = 1; i < ElementList1.Count; i++)
                //{
                //    ElementList1[i].Click();
                //}
                ElementList1[8].SendKeys(data[3]);
                ElementList1[7].SendKeys(data[4]);
                ElementList1[13].SendKeys(data[5]);
                ElementList1[15].SendKeys(data[6]);
            }
            else
            {
                ElementList1[15].Clear();
                ElementList1[15].SendKeys(data[7]);
            }
        }
    }
}
