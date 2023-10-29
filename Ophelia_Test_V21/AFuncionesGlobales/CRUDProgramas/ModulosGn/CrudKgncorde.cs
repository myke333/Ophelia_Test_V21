using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosGn
{
    class CrudKgncorde
    {
        public static void CrudKGncorde(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            List<string> lupa = new List<string> { data[0], data[1] };
            List<string> lupaEdit = new List<string> { data[0], data[2] };
            if (Tipo == 1)
            {
                PruebaCRUD.LupaDinamica2(desktopSession, lupa);
                Thread.Sleep(500);
            }
            else if (Tipo == 2)
            {
                PruebaCRUD.LupaDinamica2(desktopSession, lupaEdit);
                Thread.Sleep(500);
            }
        }
        public static void CrudKGncordeDetalle(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");

            if (Tipo == 1)
            {
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 33, 29);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(data[3]);
                Thread.Sleep(500);
            }
            else if (Tipo == 2)
            {
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 33, 29);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(data[4]);
                Thread.Sleep(500);
            }
        }
    }
}
