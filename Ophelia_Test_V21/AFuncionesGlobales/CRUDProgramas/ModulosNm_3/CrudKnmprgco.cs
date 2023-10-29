using OpenQA.Selenium.Appium.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_3
{
    class CrudKnmprgco
    {
        public static void CrudKNmprgco(WindowsDriver<WindowsElement> desktopSession, string concepto)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 50, 26);
            desktopSession.Mouse.Click(null);

            ElementList[0].SendKeys(concepto);
        }
    }
}
