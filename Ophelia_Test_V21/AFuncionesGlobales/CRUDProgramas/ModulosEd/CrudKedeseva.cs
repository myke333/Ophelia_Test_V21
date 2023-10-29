using OpenQA.Selenium.Appium.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosEd
{
    class CrudKedeseva
    {
        public static void CrudKEdeseva(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (Tipo == 1)
            {
                ElementList[3].SendKeys(data[0]);
                Thread.Sleep(500);
                ElementList[2].SendKeys(data[1]);
                Thread.Sleep(500);
            }
            else if (Tipo == 2)
            {
                ElementList[2].Clear();
                Thread.Sleep(500);
                ElementList[2].SendKeys(data[2]);
                Thread.Sleep(500);
            }
        }
        public static void CrudKEdesevaDetalle(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 33, 26);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(500);

            if (Tipo == 1)
            {
                desktopSession.Keyboard.SendKeys(data[0]);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(data[1]);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(data[2]);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(data[3]);
                Thread.Sleep(500);
            }
            else if (Tipo == 2)
            {
                for (int i = 0; i < 3; i++)
                {
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    Thread.Sleep(500);
                }
                desktopSession.Keyboard.SendKeys(data[4]);
                Thread.Sleep(500);
            }
        }
    }
}
