using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSo
{
    class CrudKsoevccs
    {
        public static void CrudKSoevccs(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            List<string> data1 = new List<string> { data[0] };
            List<string> data2 = new List<string> { data[1] };
            if (Tipo == 1)
            {
                PruebaCRUD.LupaDinamica2(desktopSession, data1);
            }
            else if (Tipo == 2)
            {
                PruebaCRUD.LupaDinamica2(desktopSession, data2);
            }
        }
        public static void CrudKSoevccsDetalle(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 66, 26);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(200);

            if (Tipo == 1)
            {
                ElementList[0].SendKeys(data[0]);
                Thread.Sleep(200);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(200);
                ElementList[0].SendKeys(data[1]);
                Thread.Sleep(200);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(200);
                ElementList[0].SendKeys(data[2]);
                Thread.Sleep(200);
            }
            else if (Tipo == 2)
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(200);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(200);
                ElementList[0].Clear();
                Thread.Sleep(200);
                ElementList[0].SendKeys(data[3]);
            }
        }
    }
}
