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
    class CrudKnmgenca
    {
        public static void CrudKNmgenca(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList1 = desktopSession.FindElementsByClassName("TDBMemo");
            List<string> lupa1 = new List<string> { data[0] };
            if (Tipo == 1)
            {
                desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Administración").Coordinates);
                Thread.Sleep(200);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                PruebaCRUD.LupaDinamica(desktopSession, lupa1);
                Thread.Sleep(500);
                ElementList1[0].SendKeys(data[1]);
                Thread.Sleep(500);
            }
            else if (Tipo == 2)
            {
                Thread.Sleep(500);
                ElementList1[0].Clear();
                Thread.Sleep(500);
                ElementList1[0].SendKeys(data[2]);
                Thread.Sleep(500);
            }
        }
    }
}
