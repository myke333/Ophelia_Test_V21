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
    class CrudKnmpretr
    {
        public static void CrudKNmpretr(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");

            if (Tipo == 1)
            {
                List<string> lupas = new List<string> { data[0], data[1], data[2], data[2], data[2] };
                PruebaCRUD.LupaDinamica2(desktopSession, lupas);
            }
            else if (Tipo == 2)
            {
                Thread.Sleep(500);
                ElementList[5].Clear();
                Thread.Sleep(500);
                ElementList[3].Clear();
            }
        }
        public static void CrudKNmpretrDetalle(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            if (Tipo == 1)
            {
                var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 66, 26);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(200);
                ElementList[0].SendKeys(data[0]);
                Thread.Sleep(200);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(200);
                ElementList[0].SendKeys(data[1]);
                Thread.Sleep(200);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(200);
                ElementList[0].SendKeys(data[2]);
            }
            else if (Tipo == 2)
            {
                var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 66, 26);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(200);
                ElementList[0].Clear();
                Thread.Sleep(200);
                ElementList[0].SendKeys(data[3]);
            }
        }
    }
}
