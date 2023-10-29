using OpenQA.Selenium.Appium.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosGn
{
    class CrudKgnarbol
    {
        public static void CrudKGnarbol(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (Tipo == 1)
            {
                Thread.Sleep(500);
                ElementList[4].SendKeys(data[0]);
                Thread.Sleep(500);
                ElementList[3].SendKeys(data[1]);
                Thread.Sleep(500);
            }
            else if (Tipo == 2)
            {
                ElementList[3].Clear();
                Thread.Sleep(500);
                ElementList[3].SendKeys(data[2]);
                Thread.Sleep(500);
            }
        }
        public static void CrudKGnarbolDetalle(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
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
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(data[2]);
                Thread.Sleep(500);
            }
        }
    }
}
