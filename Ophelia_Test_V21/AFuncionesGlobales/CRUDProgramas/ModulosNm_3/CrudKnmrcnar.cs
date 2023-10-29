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
    class CrudKnmrcnar
    {
        public static void CrudKNmrcnar(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            if (Tipo == 1)
            {
                List<string> data1 = new List<string> { data[0] };
                PruebaCRUD.LupaDinamica2(desktopSession, data1);
            }
            else if (Tipo == 2)
            {
                List<string> data1 = new List<string> { data[1] };
                PruebaCRUD.LupaDinamica2(desktopSession, data1);
            }
        }
        public static void CrudKNmrcnarDetalle(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            if (Tipo == 1)
            {
                var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 66, 26);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                ElementList[0].SendKeys(data[0]);
            }
            else if (Tipo == 2)
            {
                var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 66, 26);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                ElementList[0].SendKeys(data[1]);
            }
        }
    }
}
