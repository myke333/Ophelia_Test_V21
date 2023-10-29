using OpenQA.Selenium.Appium.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosGn
{
    class CrudKgndefsd
    {
        public static void CrudKGndefsd(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (Tipo == 1)
            {
                for (int i = 0; i < ElementList.Count; i++)
                {
                    desktopSession.Mouse.MouseMove(ElementList[i].Coordinates);
                    Thread.Sleep(2000);
                }
                //Thread.Sleep(500);
                //ElementList[5].SendKeys(data[4]);
                //Thread.Sleep(500);
                //ElementList[5].SendKeys(data[4]);
                //Thread.Sleep(500);
            }
            else if (Tipo == 2)
            {
                ElementList[5].Clear();
                Thread.Sleep(500);
                ElementList[5].SendKeys(data[5]);
                Thread.Sleep(500);
            }
        }
        public static void CrudKGndefsdDetalle(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
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
