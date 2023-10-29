using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_2
{
    class CrudKnmpreau
    {
        public static void CrudKNmpreau(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            if (Tipo == 1)
            {
                var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
                var ElementList1 = desktopSession.FindElementsByClassName("TSpinEdit");
                var ElementList2 = desktopSession.FindElementsByClassName("Edit");
                ElementList1[0].Clear();
                ElementList1[0].SendKeys(data[0]);
                Thread.Sleep(500);
                ElementList[8].SendKeys(data[1]);
                Thread.Sleep(500);
                ElementList2[2].SendKeys(data[2]);
                Thread.Sleep(500);
                ElementList[9].SendKeys(data[3]);


                Thread.Sleep(500);
                ElementList2[1].SendKeys(data[2]);
                Thread.Sleep(500);
                ElementList[10].SendKeys(data[3]);
                
                
            }
            else if (Tipo == 2)
            {
                var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
                ElementList[10].Clear();
                ElementList[10].SendKeys(data[4]);
            }
        }
    }
}
