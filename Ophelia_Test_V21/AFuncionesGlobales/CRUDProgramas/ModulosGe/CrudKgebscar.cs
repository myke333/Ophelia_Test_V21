using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosGe
{
    class CrudKgebscar
    {
        public static void CrudKGebscar(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            List<string> lupa = new List<string> { data[0], data[1], data[2], data[3] };
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (Tipo == 1)
            {
                PruebaCRUD.LupaDinamica2(desktopSession, lupa);
                Thread.Sleep(500);
                ElementList[5].SendKeys(data[4]);
                Thread.Sleep(500);
            }
            else if (Tipo == 2)
            {
                ElementList[5].Clear();
                Thread.Sleep(500);
                ElementList[5].SendKeys(data[5]);
                Thread.Sleep(500);
            }
        }
        public static void CrudKGebscarDetalle(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            if (Tipo == 1)
            {
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 33, 29);
                Thread.Sleep(500);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(data[0]);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(data[1]);
                Thread.Sleep(500);
            }
            else if (Tipo == 2)
            {
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 33, 29);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(data[2]);
                Thread.Sleep(500);
            }
        }
    }
}
