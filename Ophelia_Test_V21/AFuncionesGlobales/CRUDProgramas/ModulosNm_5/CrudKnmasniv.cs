using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_5
{
    class CrudKnmasniv
    {
        public static void CrudKNmasniv(WindowsDriver<WindowsElement> desktopSession, List<string> data)
        {
            List<string> lupa = new List<string> { data[0] };
            PruebaCRUD.LupaDinamica(desktopSession, lupa);
            Thread.Sleep(500);
        }
        public static void CrudKNmasnivDetalle(WindowsDriver<WindowsElement> desktopsession, int Tipo, List<string> data)
        {
            var ElementList = desktopsession.FindElementsByClassName("TDBGrid");

            if (Tipo == 1)
            {
                desktopsession.Mouse.MouseMove(ElementList[0].Coordinates, 33, 29);
                desktopsession.Mouse.Click(null);
                Thread.Sleep(500);
                desktopsession.Keyboard.SendKeys(data[0]);
                Thread.Sleep(500);
                desktopsession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopsession.Keyboard.SendKeys(data[1]);
                Thread.Sleep(500);
                desktopsession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopsession.Keyboard.SendKeys(data[2]);
                Thread.Sleep(500);
            }
            else if (Tipo == 2)
            {
                desktopsession.Mouse.MouseMove(ElementList[0].Coordinates, 33, 29);
                desktopsession.Mouse.Click(null);
                Thread.Sleep(500);
                desktopsession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopsession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopsession.Keyboard.SendKeys(data[3]);
                Thread.Sleep(500);
            }
        }
    }
}
