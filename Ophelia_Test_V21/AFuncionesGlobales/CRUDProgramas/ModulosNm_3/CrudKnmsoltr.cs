using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_3
{
    class CrudKnmsoltr
    {
        public static void CrudKNmsoltr(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            List<string> lupa = new List<string> { data[0] };
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList1 = desktopSession.FindElementsByClassName("TDBMemo");
            var ElementList2 = desktopSession.FindElementsByClassName("TDBCmbCodigoBox");
            if (Tipo == 1)
            {
                PruebaCRUD.LupaDinamica(desktopSession, lupa);
                Thread.Sleep(500);
                ElementList[3].SendKeys(data[1]);
                Thread.Sleep(500);
                ElementList[2].SendKeys(data[2]);
                Thread.Sleep(500);
                ElementList[4].SendKeys(data[3]);
                Thread.Sleep(500);
                desktopSession.Mouse.MouseMove(ElementList2[0].Coordinates);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                ElementList2[0].SendKeys(data[4]);
                Thread.Sleep(500);
                ElementList2[0].SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(500);
                ElementList1[0].SendKeys(data[5]);
                Thread.Sleep(500);
            }
            else if (Tipo == 2)
            {
                ElementList1[0].Clear();
                Thread.Sleep(500);
                ElementList1[0].SendKeys(data[6]);
                Thread.Sleep(500);
            }
        }
    }
}
