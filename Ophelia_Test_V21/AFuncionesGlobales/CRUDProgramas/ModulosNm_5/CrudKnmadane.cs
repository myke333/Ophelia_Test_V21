using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_5
{
    class CrudKnmadane
    {
        public static void CrudKNmadane(WindowsDriver<WindowsElement> desktopSession, List<string> data)
        {
            List<string> lupa = new List<string> { data[0] };
            var ElementList = desktopSession.FindElementsByClassName("TEdit");
            PruebaCRUD.LupaDinamica2(desktopSession, lupa);
            Thread.Sleep(500);
            Thread.Sleep(500);
            ElementList[7].SendKeys(data[1]);
            Thread.Sleep(500);
            ElementList[6].SendKeys(data[2]);
            Thread.Sleep(500);
        }
    }
}
