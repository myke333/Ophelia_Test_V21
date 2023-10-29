using OpenQA.Selenium.Appium.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_5
{
    class CrudKnmappin
    {
        public static void CrudKNmappin(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList1 = desktopSession.FindElementsByClassName("TDBComboBox");
            if (Tipo == 1)
            {
                Thread.Sleep(500);
                ElementList[12].SendKeys(data[0]);
                Thread.Sleep(500);
                desktopSession.Mouse.MouseMove(ElementList1[0].Coordinates);
                Thread.Sleep(500);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("E - NORMAL").Coordinates);
                Thread.Sleep(500);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                ElementList[10].SendKeys(data[1]);
                Thread.Sleep(500);
            }
            else if (Tipo == 2)
            {
                ElementList[10].Clear();
                Thread.Sleep(500);
                ElementList[10].SendKeys(data[2]);
                Thread.Sleep(500);
            }
        }
    }
}
