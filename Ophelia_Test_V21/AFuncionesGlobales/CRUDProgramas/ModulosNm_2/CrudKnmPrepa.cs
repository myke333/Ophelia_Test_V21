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
    class CrudKnmprepa
    {
        public static void CrudKNmprepa(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TSpinEdit");
            var ElementList1 = desktopSession.FindElementsByClassName("TDBEdit");
            
            if (Tipo == 1)
            {
                List<string> lupas = new List<string> { data[3], data[3], data[3], data[3], data[3], data[3], data[3], data[3], data[3], data[3] };
                ElementList[0].Clear();
                ElementList[0].SendKeys(data[0]);
                desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Cargos").Coordinates);
                desktopSession.Mouse.Click(null);
                ElementList1[13].SendKeys(data[1]);
                ElementList1[24].SendKeys(data[2]);
                PruebaCRUD.LupaDinamica2(desktopSession, lupas);
            }
            else if (Tipo == 2)
            {
                ElementList1[13].Clear();
                ElementList1[13].SendKeys(data[4]);
            }
        }
    }
}
